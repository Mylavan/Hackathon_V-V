from enum import Enum
from openpyxl import Workbook
import re
import fitz  # PyMuPDF
import pdfplumber
#**************************************************************FIXED KEYWORD********************************************************
End_of_Document_Keyword = "End of Document"
 
class Status(Enum):
    START = 0
    ALLNONE = 1
    INTIALSNOTMATHCING = 2
    SUCCESS = 3 
#**************************************************************FIXED KEYWORD********************************************************

def is_diagonal_annot(annot, tolerance=0.9):
    info = annot.info
    rect = annot.rect   
    if "vertices" in info:
        vertices = info["vertices"]
        if len(vertices) == 2:
            x0, y0 = vertices[0]
            x1, y1 = vertices[1]
            dx = x1 - x0
            dy = y1 - y0
            if dx == 0:
                return False
            slope = dy / dx
            return abs(abs(slope) - 1) < tolerance
    # fallback to rect size
    width = rect.width
    height = rect.height
    if width == 0:
        return False
    slope = height / width
    return abs(abs(slope) - 1) < tolerance

def extract_strike_y_ranges(pdf_path):
    doc = fitz.open(pdf_path)
    strike_ranges = {}

    for page_num, page in enumerate(doc, start=1):
        diagonal_annots = []

        for annot in page.annots():
            if annot.type[0] == 3:  # line annotation
                if is_diagonal_annot(annot):
                    diagonal_annots.append(annot)

        if not diagonal_annots:
            continue  # Skip if no diagonal lines

        # Collect Y coordinates of all diagonal annotations
        strike_y_values = []
        for annot in diagonal_annots:
            info = annot.info
            if "vertices" in info:
                y_coords = [v[1] for v in info["vertices"]]
                strike_y_values.extend(y_coords)
            else:
                strike_y_values.extend([annot.rect.y0, annot.rect.y1])

        min_strike_y = min(strike_y_values)
        max_strike_y = max(strike_y_values)

        strike_ranges[page_num-1] = (int(min_strike_y), int(max_strike_y))
        #print(f"🔵 Strike Y range: {min_strike_y:.2f} - {max_strike_y:.2f}")

        # Extract text blocks
        text_blocks = page.get_text("dict")["blocks"]
        all_lines = []

        for block in text_blocks:
            if block["type"] != 0:
                continue

            for line in block["lines"]:
                y_top = line["bbox"][1]
                line_text = " ".join(span["text"] for span in line["spans"]).strip()
                if line_text:
                    all_lines.append((y_top, line_text))

        # Sort lines from top (smallest y) to bottom (largest y)
        all_lines.sort(key=lambda x: x[0])

        # Assign visual line numbers
        numbered_lines = [(i + 1, y, text) for i, (y, text) in enumerate(all_lines)]

        # Split above and below strike
        above_lines = [(lnum, text) for lnum, y, text in numbered_lines if y < min_strike_y]
        below_lines = [(lnum, text) for lnum, y, text in numbered_lines if y > max_strike_y]

        # print("\n⬆️ Lines ABOVE strike:")
        # for lnum, text in above_lines:
        #     print(f"Line {lnum}: {text}")

        # print("\n⬇️ Lines BELOW strike:")
        # for lnum, text in below_lines:
        #     print(f"Line {lnum}: {text}")
    return strike_ranges


def is_empty(value):
    return value in [None, "", "None", "______", "___________"]

def extract_text_and_form_fields_fromStrikedLines_old(pdf_path, Intials, y_exclude_dict):
    Intials_Issues = []
    FormField_Issues = []
    Intials_Mismatches_Footer = []
    Date_Mismatches_Footer = []
    IssueField_Pagenumber_Holder = []
    date = None

    doc = fitz.open(pdf_path)

    for page_num, page in enumerate(doc):
        print(f"\n📄 Page {page_num + 1}")
        print("-" * 40)

        y_min, y_max = y_exclude_dict.get(page_num, (None, None))
        IssueField_Pagenumber_Holder_temp = []

        blocks = page.get_text("dict")["blocks"]
        size_threshold = 12.0
        line_counter = 0  # <-- track line number

        for b in blocks:
            if "lines" in b:
                for l in b["lines"]:
                    line_counter += 1
                    for s in l["spans"]:
                        text = s["text"].strip()
                        size = round(s["size"], 1)

                        if text and size == size_threshold:
                            print(f"🔹 Heading (line {line_counter}) on page {page_num+1}: {text} (size {size})")
                            IssueField_Pagenumber_Holder_temp.clear()
                            IssueField_Pagenumber_Holder_temp.append((page_num + 1, line_counter, text))

        # ==== FORM FIELD EXTRACTION ====
        person = {}
        widgets = page.widgets()
        if widgets:
            for idx, widget in enumerate(widgets, start=1):  # enumerate widgets like line numbers
                top_y = widget.rect.y0
                if y_min is not None and y_max is not None and y_min <= top_y <= y_max:
                    continue

                field_name = widget.field_name
                field_value = widget.field_value
                if is_empty(field_value):
                    field_value = "None"
                person[field_name] = field_value

                print(f"📝 Widget (line {idx}) on page {page_num+1}: {field_name} = {field_value}")

            # sort & process fields
            def sort_key(item):
                match = re.search(r'_(\d+)$', item[0])
                return int(match.group(1)) if match else float('inf')

            person = dict(sorted(person.items(), key=sort_key))

            for key, value in person.items():
                if "Date" in key:
                    if (page_num+1) != 1:
                        if (date != value):
                            Date_Mismatches_Footer.append((page_num + 1))
                    date = value
                elif "Initials" in key:
                    if (value != Intials):
                        Intials_Mismatches_Footer.append((page_num + 1))
                else:
                    suffixes = set()
                    for k in person:
                        if any(prefix in k for prefix in ['Pass_', 'Fail_', 'Issue_']):
                            parts = k.split("_")
                            if len(parts) == 2 and parts[1].isdigit():
                                suffixes.add(parts[1])

                    for suffix in sorted(suffixes, key=int):
                        pass_key = f"Pass_{suffix}"
                        fail_key = f"Fail_{suffix}"
                        issue_key = f"Issue_{suffix}"

                        pass_val = person.get(pass_key)
                        fail_val = person.get(fail_key)
                        issue_val = person.get(issue_key)

                        if is_empty(pass_val) and is_empty(fail_val) and is_empty(issue_val):
                            FormField_Issues.append(page_num + 1)
                        elif (pass_val == "None" and fail_val == "None"):
                            FormField_Issues.append(page_num + 1)
                        elif pass_val != "None":
                            if pass_val != Intials:
                                Intials_Issues.append(page_num + 1)
                                print(f"  ❌ Pass value mismatch on page {page_num + 1}: {pass_val} != {Intials}")
                        elif fail_val != "None":
                            if fail_val != Intials:
                                Intials_Issues.append(page_num + 1)
                            elif issue_val == "None":
                                Intials_Issues.append(page_num + 1)
                            else:
                                print(f"  ✅ Fail value and issue are valid.")
                                IssueField_Pagenumber_Holder.append(IssueField_Pagenumber_Holder_temp)

    FormField_Issues = list(dict.fromkeys(FormField_Issues))
    Intials_Issues = list(dict.fromkeys(Intials_Issues))
    IssueField_Pagenumber_Holder_temp.clear()

    return FormField_Issues, Intials_Issues, date, Intials_Mismatches_Footer, Date_Mismatches_Footer, IssueField_Pagenumber_Holder

def remove_duplicates_from_list(input_list):
    unique_list = []
    for item in input_list:
        if item not in unique_list:
            unique_list.append(item)
    return unique_list


def clean_heading_text(heading_text):
    """Extract only the section number like 7.3.1, 7.4, etc."""
    match = re.match(r"^\d+(\.\d+)*", heading_text)
    return match.group(0) if match else heading_text

def map_issues_to_headings(headings, issues):
    result = []
    for issue_page, issue_line, issue_text in issues:
        issue_line = int(issue_line)
        best_heading = None
        for h_page, h_line, h_text in headings:
            if h_page < issue_page or (h_page == issue_page and h_line <= issue_line):
                best_heading = clean_heading_text(h_text)
        if best_heading:
            result.append((best_heading, issue_text))
    return result





def extract_text_and_form_fields_fromStrikedLines(pdf_path, Intials, y_exclude_dict):
    Intials_Issues = []
    FormField_Issues = []
    Intials_Mismatches_Footer = []
    Date_Mismatches_Footer = []
    Heading_Line_Holder = []
    Issue_Line_page_Holder=[]
    date = None

    doc = fitz.open(pdf_path)

    for page_num, page in enumerate(doc):
        print(f"\n📄 Page {page_num + 1}")
        print("-" * 40)

        y_min, y_max = y_exclude_dict.get(page_num, (None, None))

        blocks = page.get_text("dict")["blocks"]
        size_threshold = 12.0
        line_counter = 0  # <-- track line number
        line_map = []     # <-- stores (line_num, y0, text) for mapping widgets

        # ==== TEXT (Headings) EXTRACTION ====
        for b in blocks:
            if "lines" in b:
                for l in b["lines"]:
                    line_text = " ".join([s["text"].strip() for s in l["spans"] if s["text"].strip()])
                    if not line_text:
                        continue

                    line_counter += 1
                    y0 = l["bbox"][1]  # Y-coordinate of line
                    line_map.append((line_counter, y0, line_text))

                    for s in l["spans"]:
                        text = s["text"].strip()
                        size = round(s["size"], 1)

                        if text and size == size_threshold:
                            #print(f"🔹 Heading (line {line_counter}) on page {page_num+1}: {text} (size {size})")
                            for char in text:
                                if char.isdigit():                                 
                                    Heading_Line_Holder.append((page_num+1,line_counter, text))

        # ==== FORM FIELD EXTRACTION ====
        person = {}
        widgets = page.widgets()
        if widgets:
            for widget in widgets:
                top_y = widget.rect.y0
                if y_min is not None and y_max is not None and y_min <= top_y <= y_max:
                    continue

                field_name = widget.field_name               
                field_value = widget.field_value
                if is_empty(field_value):
                    field_value = "None"

                # 🔗 Find nearest line above widget
                mapped_line = None
                for ln, y0, txt in sorted(line_map, key=lambda x: -x[1]):  # sort descending
                    if y0 <= top_y:
                        mapped_line = ln
                        break
                field_name=re.sub(r'_(\d+)$', f'_{mapped_line}', field_name) #replace the random appended with line number            
                #print("new field name", field_name)
                person[field_name] = field_value                            
                #print(f"📝 Widget (line {mapped_line}) on page {page_num+1}: {field_name} = {field_value}")

            # ==== SORT & VALIDATE ====
            def sort_key(item):
                match = re.search(r'_(\d+)$', item[0])
                return int(match.group(1)) if match else float('inf')

            person = dict(sorted(person.items(), key=sort_key))

            for key, value in person.items():
                if "Date" in key:
                    if (page_num+1) != 1:
                        if (date != value):
                            Date_Mismatches_Footer.append((page_num + 1))
                    date = value
                elif "Initials" in key:
                    if (value != Intials):
                        Intials_Mismatches_Footer.append((page_num + 1))
                else:
                    suffixes = set()
                    for k in person:
                        if any(prefix in k for prefix in ['Pass_', 'Fail_', 'Issue_']):
                            parts = k.split("_")
                            if len(parts) == 2 and parts[1].isdigit():
                                suffixes.add(parts[1])

                    for suffix in sorted(suffixes, key=int):
                        pass_key = f"Pass_{suffix}"
                        fail_key = f"Fail_{suffix}"
                        issue_key = f"Issue_{suffix}"

                        pass_val = person.get(pass_key)
                        fail_val = person.get(fail_key)
                        issue_val = person.get(issue_key)

                        if is_empty(pass_val) and is_empty(fail_val) and is_empty(issue_val):
                            FormField_Issues.append(page_num + 1)
                        elif (pass_val == "None" and fail_val == "None"):
                            FormField_Issues.append(page_num + 1)
                        elif pass_val != "None":
                            if pass_val != Intials:
                                Intials_Issues.append(page_num + 1)
                                #print(f"  ❌ Pass value mismatch on page {page_num + 1}: {pass_val} != {Intials}")
                        elif fail_val != "None":
                            if fail_val != Intials:
                                Intials_Issues.append(page_num + 1)
                            elif issue_val == "None":
                                Intials_Issues.append(page_num + 1)
                            else:
                                #print(f"  ✅ Fail value and issue are valid.")                               
                                Line_number=pass_key.split("_")[-1]
                                #print(f"  📝 Issue found on page {page_num + 1}, line {Line_number}: {issue_val}")
                                Issue_Line_page_Holder.append((page_num + 1, Line_number, issue_val))

    FormField_Issues = list(dict.fromkeys(FormField_Issues))
    Intials_Issues = list(dict.fromkeys(Intials_Issues))
    Cleaned_Heading_Line_Holder = list(dict.fromkeys(Heading_Line_Holder))  #removing the duplicates from the list
    Cleaned_Issue_Line_page_Holder = list(dict.fromkeys(Issue_Line_page_Holder)) #removing the duplicates from the list
    Issue_Heading = map_issues_to_headings(Cleaned_Heading_Line_Holder, Cleaned_Issue_Line_page_Holder) #mapping the issue and heading based on page number and line number 
    return FormField_Issues, Intials_Issues, date, Intials_Mismatches_Footer, Date_Mismatches_Footer,  Issue_Heading

def Endkeyword_in_pdf(pdf_path: str, keyword: str) -> None:
    """
    Searches for a keyword in a PDF and prints the lines containing the keyword with page numbers.

    Args:
        pdf_path (str): Path to the PDF file.
        keyword (str): Keyword to search for (case-sensitive).
    """
    found = False  # Flag to track if keyword was found

    with pdfplumber.open(pdf_path) as pdf:
        #print("📄 PDF Line Extraction with Keyword Filter\n" + "=" * 50)

        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()

            if text:
                lines = text.splitlines()
                for line in lines:
                    if keyword in line:
                        #print(f"[Page {page_num + 1}] {line}")
                        found = True

    return found

def extract_footer_from_first_page(pdf_path: str) -> str:
    """
    Extracts footer text from the first page of a PDF (bottom 9% of the page).

    Args:
        pdf_path (str): Path to the PDF file.

    Returns:
        str: Extracted footer text, or an empty string if no footer is found.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]  # Only the first page
        page_height = page.height
        footer_threshold = page_height * 0.91  # Bottom 9% of the page

        footer_text = [
            word["text"] for word in page.extract_words()
            if float(word["top"]) >= footer_threshold
        ]
    #print("Footer text extracted from first page:", " ".join(footer_text))
    footer_str = " ".join(footer_text)

    # 1. Extract till "Baseline"
    Test_Procedure_Name = footer_str.split("Baseline")[0].strip()

    # 2. Extract between "Baseline" and "Initials"
    match = re.search(r'Baseline: (.*?) Initials:', footer_str)
    Baseline = match.group(1).strip() if match else None
    return  Test_Procedure_Name, Baseline



def Test_Procedure_Entry(pdf_path,Initials):  
    CurrentStatus = 0
    Intials_Issues=[]
    FormField_Issues=[]
    Intials_Mismatches_Footer=[]
    Date_Mismatches_Footer=[]
    Issue_Heading=[]
    EndKeywordPresent=0
    Test_Procedure_Name= ""
    Baseline=0
    Test_Procedure_data = {
        "Filename": pdf_path.split("\\")[-1],
        "FileType": "Test Procedure",
        "Initials": Initials,
        "Start Date": 0,       
        "Intials_Issues": [],
        "FormField_Issues": [],
        "Intials_Mismatches_Footer": [],
        "Date_Mismatches_Footer": [],
        "Issue_Heading": [],
        "EndKeywordPresent": 0,
        "Test Procedure Name": "",
        "Baseline": 0
    }
    StrikedLines_StartY_EndY=extract_strike_y_ranges(pdf_path)   
    Test_Procedure_data["FormField_Issues"], Test_Procedure_data["Intials_Issues"], Test_Procedure_data["Start Date"], Test_Procedure_data["Intials_Mismatches_Footer"], Test_Procedure_data["Date_Mismatches_Footer"], Test_Procedure_data["Issue_Heading"] = extract_text_and_form_fields_fromStrikedLines(  pdf_path, Test_Procedure_data["Initials"],StrikedLines_StartY_EndY )


    #print("FormField_Issues: ", Test_Procedure_data["FormField_Issues"])   
 
    Test_Procedure_data["EndKeywordPresent"] = Endkeyword_in_pdf(pdf_path, End_of_Document_Keyword)
    
    #print("EndKeywordPresent: ", Test_Procedure_data["EndKeywordPresent"])

    Test_Procedure_data["Test Procedure Name"], Test_Procedure_data["Baseline"] = extract_footer_from_first_page(pdf_path)


    #print("Test_Procedure_Name:", Test_Procedure_data["Test_Procedure_Name"])
    #print("Baseline:", Test_Procedure_data["Baseline"])

    return Test_Procedure_data
  

# if __name__ == "__main__":
#     pdf_path =r"C:\DEMO_TEST\D001929432 1\crossed_Test Data Sheet.pdf"
#     Test_Procedure_Entry(pdf_path,"NR")
   
#     #extract_text_and_form_fields_org(r"C:\DEMO_TEST\D001929432 1\crossed_Test Data Sheet.pdf","NR")

