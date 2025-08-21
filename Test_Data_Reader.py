import pdfplumber
import os
import xml.etree.ElementTree as ET
import re
from pypdf import PdfReader
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import os



#Pdf to XML conversion function
def extract_tables_to_xml(pdf_path, xml_output_path):
    if os.path.exists(xml_output_path):
        print(f"File already exists: {xml_output_path}")
        return

    root = ET.Element("Tables")

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for table_index, table in enumerate(tables):
                table_el = ET.SubElement(root, "Table", page=str(i), index=str(table_index))
                for row in table:
                    row_el = ET.SubElement(table_el, "Row")
                    for cell in row:
                        cell_text = cell.strip() if cell else ""
                        ET.SubElement(row_el, "Cell").text = cell_text

    tree = ET.ElementTree(root)
    tree.write(xml_output_path, encoding='utf-8', xml_declaration=True)
    print(f"Extracted tables written to: {xml_output_path}")

def clean_text(text: str) -> str:
    """Remove known label patterns and return cleaned text."""
    if not text:
        return ""
    text = text.strip()

    # Remove content inside parentheses e.g. (printed), (signature)
    text = re.sub(r"\(.*?\)", "", text)

    # Strip labels like 'Test Folder Number:', 'Start Date:', etc.
    text = re.sub(r"^[A-Za-z ]*:", "", text)

    return text.strip()


def row_has_meaningful_content(row) -> bool:
    """Check if row has at least one cell with meaningful data."""
    for cell in row.findall("Cell"):
        if cell.text and clean_text(cell.text):
            return True
    return False


def table_has_content(xml_file, page, index) -> bool:
    """Check if table has meaningful content rows."""
    tree = ET.parse(xml_file)
    root = tree.getroot()

    for table in root.findall(".//Table"):
        if table.get("page") == str(page) and table.get("index") == str(index):
            for row in table.findall("Row"):
                if row_has_meaningful_content(row):
                    return True
            return False
    return False


keys_from_table = [
    "Test Folder Number",
    "Start Date",
    "Stop Date",
    "Actual Hours",
    "Tester Name",
    "Initials",
    "Test Procedure Name",
    "Test Procedure View",
    "Baseline",
    "Location",
    "Product",
    "Subproduct",
    "Configuration",
    "System Software Version Number",
    "Project Name",
 
]
# Create lowercase key mapping for comparison
keys_from_table_map = {key: "Not found" for key in keys_from_table}
found_key_values = {}

## From Test Folder Number till Project Name, all keys are extracted [First 2 tables extracted from the XML]
def First_table_Software_Configuration(xml_path):

    
    tree = ET.parse(xml_path)
    root = tree.getroot()
  
    for table_index, table in enumerate(root.findall("Table")):
        page = table.attrib.get("page", "N/A")
        index = table.attrib.get("index", str(table_index))

        for row in table.findall("Row"):
            cells = [cell.text.strip() if cell.text else "" for cell in row.findall("Cell")]
           # print(" | ".join(cells))

            # Check for key presence in any cell
            for i, cell_text in enumerate(cells):  # avoid last cell
                cell_text = cell_text.strip()
                inner_text = cell_text.strip().strip("'")  # remove leading/trailing quotes
                parts = inner_text.partition(":")
                key = parts[0].strip()
                value = parts[2].split("'")[0].strip()
               
         
                if cell_text.strip() and inner_text.strip():

                 if value.strip():
                    if key in keys_from_table_map:             
                       
                        if(key =="Tester Name"): #need to remove (Printed) from the value 
                            value = re.sub(r'\([^)]*\)', '', value).strip()
                        
                        found_key_values[key] = value
                 else:
                       if( "initials" in key):  # extract intials
                           value = re.sub(r'\([^)]*\)', '', key).strip()
                           found_key_values["Initials"] = value


    # Print extracted key-value pairs
    #print("\n🔑 Extracted Key-Value Pairs:")
    # for key in keys_from_table:
    #     value = found_key_values.get(key, "<Not Found>")
    #     print(f"{key}: {value}")
    return found_key_values

def Specific_table_extractor(xml_path, target_page, target_index, target_row_index):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    Output = []
    found = False
    stop_phrase = "If more space required"

    for table_index, table in enumerate(root.findall("Table")):
        try:
            page = int(table.attrib.get("page", "-1"))
            index = int(table.attrib.get("index", str(table_index)))
        except ValueError:
            continue  # skip malformed attributes

        if page == target_page and index == target_index:
            #print(f"\n📄 Table {index} (from page {page}), Rows > {target_row_index}:")
            #print("-" * 60)

            rows = table.findall("Row")
            if target_row_index + 1 >= len(rows):
                print(f"⚠️ No rows found after row {target_row_index}.")
            else:
                for i in range(target_row_index + 1, len(rows)):
                    row = rows[i]
                    cells = [cell.text.strip() if cell.text else "" for cell in row.findall("Cell")]

                    # Check for stop phrase in any cell
                    if any(stop_phrase in cell for cell in cells):
                        #print("🛑 Stop phrase found. Stopping before printing this row.")
                        break
                   
                    row_text = " | ".join(cells)
                    Output.append(row_text)
                    #print(f"Row {i+1}: {row_text}")

            #print("-" * 60)
            found = True
            break  # Stop after finding the specific table

    if not found:
        print(f"No table found for page {target_page} and index {target_index}")
    return Output

def CleanDatafromtable(raw_data):
   cleaned = []
   for item in raw_data:
       parts = [p.strip() for p in item.split('|')]
       # keep non-empty, not 'S', and not 'CNR'
       filtered = [p for p in parts if p and p != 'S'  and p != 'N/A']
       cleaned.append(filtered)
   return cleaned


def Table_Extractor_Entry(pdf_path, xml_path):
    extract_tables_to_xml(pdf_path, xml_path)
    FirstTableEmpty = table_has_content(xml_path, page=1, index=0)
    print("first table emprty ::")
    print(FirstTableEmpty)
    Increase_the_Table_number = 1 if FirstTableEmpty else 0
    print("Increase_the_Table_number")
    print(Increase_the_Table_number)
    First_Table = First_table_Software_Configuration(xml_path)
    SystemUsedRaw = Specific_table_extractor(xml_path, 1, 3-Increase_the_Table_number, 3)  # Extract first table from first page ::::: systems_used
    FailureSummaryRaw = Specific_table_extractor(xml_path, 1, 4-Increase_the_Table_number, 3)  # Extract second table from first page :::::Failure_Summary
    TestEquipmentRaw1 = Specific_table_extractor(xml_path, 1, 5-Increase_the_Table_number, 6)  # Extract second table from first page :::::Test_Equipment
    TestEquipmentRaw2 = Specific_table_extractor(xml_path, 2, 1, -1)  # Extract first table from second page :::::Test_Equipment
    TestEquipmentRaw= TestEquipmentRaw1 + TestEquipmentRaw2
    SystemUsedRaw = CleanDatafromtable(SystemUsedRaw)
    TestEquipmentRaw = CleanDatafromtable(TestEquipmentRaw)
    FailureSummaryRaw = CleanDatafromtable(FailureSummaryRaw)

    return (First_Table, SystemUsedRaw, FailureSummaryRaw, TestEquipmentRaw)
def truncate_line_at_keyword(line, keyword):
    if keyword in line:
        return line.split(keyword)[0].strip()
    return line.strip()

def extract_keys_from_pdf(pdf_path, keys_to_extract):
    reader = PdfReader(pdf_path)
    extracted_data = {key: "Not found" for key in keys_to_extract}

    lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            lines.extend(text.split('\n'))

    i = 0
    while i < len(lines):
        line = lines[i].strip()     

        # === DEFAULT EXTRACTION LOGIC ===
        for key in keys_to_extract:
            if key in line:
                parts = line.split(key)
                value = parts[1].strip().lstrip(":").strip() if len(parts) > 1 else ""
                if value == "" and i + 1 < len(lines):
                    value += " " + lines[i + 1].strip()

                extracted_data[key] = value.strip()

        i += 1

    return extracted_data

def Test_Data_Extraction( pdf_path):
    keys = [
        
        "PDM Doc. ID",
        "Document ID:",
        "Document Version:",
        "ARIS Template ID:",
        "ARIS Template Version:",
    ]

    result = extract_keys_from_pdf(pdf_path, keys)
    Filename = {"Filename": os.path.basename(pdf_path)}
    FileType = {"FileType": "TEST DATA"}
    result = {**Filename, **FileType, **result}

    #print("✅ Extracted Test Data Results:")
    # for k, v in result.items():
    #     print(f"{k}: {v}")

    return result



# Example usage
def Test_Data_Reader_Entry(pdf_path):
    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    print(pdf_path)
    xml_file = "output_tables.xml"
    First_table,SystemUsedRaw,FailureSummaryRaw,TestEquipmentRaw, = Table_Extractor_Entry(pdf_path, xml_file)
    Text_Extract = Test_Data_Extraction(pdf_path)

    if os.path.exists(xml_file):
     os.remove(xml_file)  #Remove the XML file after using  them to infernece 

    Final_Extract=First_table | Text_Extract
    # print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    # print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    # print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    # print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    # print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    
    # for k, v in Final_Extract.items():
    #     print(f"{k}: {v}")
    # print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    for item in SystemUsedRaw:
        print("System Used:", item)
    print("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    for item in TestEquipmentRaw:
        print("Test Equipment:", item)
    print("################################################################")
    for item in FailureSummaryRaw:
        print("Failure Summary:", item)
    return(Final_Extract, SystemUsedRaw, FailureSummaryRaw, TestEquipmentRaw)
