#TODO:Known issue: Cant read content control fields from word.
# Solution: Remove content control from word, or get Qualification id from document name.
#TODO: Newmodule, to create folder for each test read in file
#TODO: Based on test name and chapter try to get pages from pdf
#TODO: ADD data read from TO to a word template called Tracking form.

from path_selector import path_selector
from TOData import ReadTOData
import database

# Check if path to documents are available, if not found request them from user.
if "doc_path" not in database.load_database():
    path_selector()



# Creates a new object from the class ReadTOData
new_project = ReadTOData()

# Object current_data contain all necessary data from a Test Order and project tracking
current_data = new_project.get_data()

# Create the folder structure based on the infromation from TO
new_project.create_folder_structure(current_data)

# Write current data to json database and returns an object with all projects
database.database(current_data)

# print(current_data)
# Add values from current project to test tracking excel
new_project.complete_test_tracking(current_data["ProjectID"])

# Function to create Test Report, input project id and number of test report.
from create_test_report import create_report
create_report("R02074", "6")