from Test_Data_Reader import Test_Data_Reader_Entry
from Test_Procedure_reader import Test_Procedure_Entry
from Excel_output_writter import Excel_output_Entry
import os
import sys

# Main entry point to process test data and procedure files from a folder
def Main_Helper_Entry(folder_path):
    print(f"✅ main_func started with folder_path: {folder_path}")
    keyword = "Test Data Sheet"

    # Get all PDF files in the folder
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    print(f"📄 PDFs found in folder: {pdf_files}")

    # Separate test data files and other PDFs
    test_data_files = [f for f in pdf_files if keyword.lower() in f.lower()]
    other_files = [f for f in pdf_files if f not in test_data_files]

    if not test_data_files:
        print(f"❌ No test data file found with keyword: {keyword}")
        sys.exit(1)

    # Process the first test data file
    test_data_path = os.path.join(folder_path, test_data_files[0])
    print(f"🔍 Processing Test Data File: {test_data_files[0]}")
    TestData_Results, SystemUsedRaw, FailureSummaryRaw, TestEquipmentRaw = Test_Data_Reader_Entry(test_data_path)

    # Process each test procedure file and keep the last result for Excel output
    print("###################################################################################")
    TestProcedure_Results = None
    for file in other_files:
        print(f"🔍 Processing Test Procedure File: {file}")
        test_procedure_path = os.path.join(folder_path, file)
        TestProcedure_Results = Test_Procedure_Entry(test_procedure_path, TestData_Results.get("Initials", ""))

    # Print test data results for verification
    for key, value in TestData_Results.items():
        print(f"{key}: {value}")

    # Write all results to Excel
    Excel_output_Entry(TestData_Results, TestProcedure_Results, SystemUsedRaw, FailureSummaryRaw, TestEquipmentRaw,folder_path)
    print("Code completed successfully.")