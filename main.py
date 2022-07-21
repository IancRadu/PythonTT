# TODO:Known issue: Cant read content control fields from word.
# Solution: Remove content control from word, or get Qualification id from document name.
# TODO: Newmodule, to create folder for each test read in file
# TODO: Based on test name and chapter try to get pages from pdf
# TODO: ADD data read from TO to a word template called Tracking form.

from path_selector import path_selector
from TOData import ReadTOData
import database
import get_qp_pdf_info

# Check if path to documents are available, if not found request them from user.
if "doc_path" not in database.load_database():
    path_selector()

user_answer = int(
    input("What do you want to do? \n 1. Add information from test order to test tracking, create project "
          "folders and get the snipping from QP. \n 2. Create test reports from a previous imported project.\n 3. Exit.\n"))
print(user_answer)
while user_answer != 3:
    if user_answer == 1:
        print(
            "Please select the Test Order. This will automatically add your test order information to test tracking and it will "
            "create folders for each test.")
        # Creates a new object from the class ReadTOData
        new_project = ReadTOData()

        # Object current_data contain all necessary data from a Test Order and project tracking
        current_data = new_project.get_data()
        print("Wait for the program to make the necessary changes")
        # Create the folder structure based on the information from TO
        new_project.create_folder_structure(current_data)

        # Get information from pdf document QP and add it to current data
        get_qp_pdf_info.get_page_number(current_data)

        # Export pdf pages to the folder
        get_qp_pdf_info.output_pdf_page(current_data)

        # Write current data to json database and returns an object with all projects
        database.database(current_data)

        # print(current_data)
        # Add values from current project to test tracking excel
        new_project.complete_test_tracking(current_data["ProjectID"])
        user_answer = int(
            input("What do you want to do next? \n 1. Add information from test order to test tracking, create project "
                  "folders and get the snipping from QP. \n 2. Create test reports from a previous imported project.\n 3. Exit.\n"))
    elif user_answer == 2:
        # Function to create Test Report, input project id and number of test report.
        from create_test_report import create_report

        project_id = str(
            input("To create a test report first you need to write the Project ID. example: R02110\n")).strip().upper()
        if len(project_id) != 6:
            project_id = input("Write a valid project ID number\n").strip().upper()
        test_number = input("Secondly you need to write the test number. example: 2\n").strip()
        create_report(project_id, test_number)
        user_answer_report = 'yes'
        while user_answer_report == 'yes':
            user_answer_report = str(input("Do you want to create a new report? Write yes or no.\n")).lower().strip()
            test_number = str(input("Write the test number to create another test report. example: 2\n").strip())
            create_report(project_id, test_number)
        # create_report("R02074", "6") # used for testing
        user_answer = int(
            input("What do you want to do next? \n 1. Add information from test order to test tracking, create project "
                  "folders and get the snipping from QP. \n 2. Create test reports from a previous imported project.\n 3. Exit.\n"))