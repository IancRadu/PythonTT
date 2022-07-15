import pathlib
from tkinter import filedialog
import time
import database
# Function used to get path to documents and explain to the user how to do so
def path_selector():
    print("Hello, first time users need to select the path to several documents before using the software."
          "\n")
    print(
        "First document: Laboratory_equipment_configuration_update_RD3.xlsx. Found in 300_QL -> Equipment Inventory folder."
        "\n")
    print(
        "Second document: CAP1013201-F13 Equipment Planning.xlsx. Go on sharepoint where file is located and select this \n"
        "document, then click Add shortcut to OneDrive. Wait for the file to appear on your personal OneDrive and select it.\n"
        "")
    print(
        "Third document: CA 1014449-A13 QL SBZ REL Projects and Tests Tracking Overview.xlsx. Go on sharepoint where file is located\n"
        "and select this document, then click Add shortcut to OneDrive. Wait for the file to appear on your personal OneDrive\n"
        "and select it.\n"
        "")
    answer = input(
        "Write yes if you managed to do the above steps to begin selecting the paths for the files:\n").lower()
    if answer == 'yes':
        time.sleep(2)
        print("Select first document: Laboratory_equipment_configuration_update_RD3.xlsx")
        eq_configuration = pathlib.PureWindowsPath(filedialog.askopenfile().name)
        print(eq_configuration)
        print(str(eq_configuration))
        time.sleep(2)
        print("Select second document: CAP1013201-F13 Equipment Planning.xlsx")
        planning = pathlib.PureWindowsPath(filedialog.askopenfile().name)
        time.sleep(2)
        print("Select third document: CA 1014449-A13 QL SBZ REL Projects and Tests Tracking Overview.xlsx")
        output_location = pathlib.PureWindowsPath(filedialog.askopenfile().name)
        regular_path = {'doc_path': {'eq_configuration': str(eq_configuration),
                                     'planning': str(planning),
                                     'output_location': str(output_location)}}
        database.upload_database(regular_path)
    else:
        print("You had one job, close and open the app again")

