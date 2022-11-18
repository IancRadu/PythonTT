import pathlib

from docxtpl import DocxTemplate, RichText, InlineImage
import pandas as pd
import datetime
import database
from docx.shared import Mm
from test_method import test_method
# Path to the template used
tpl = DocxTemplate('./Template/Test tracking form.docx')

# Assign path to general files to variable
planning = database.load_database()['doc_path']['planning']
eq_configuration = database.load_database()['doc_path']['eq_configuration']

#
def get_chamber_data(data, planning, report_number):
    # Check equipment planning and return chamber name on which test is planned
    read_data = pd.read_excel(planning, "ENV2020", header=5)
    all_data = read_data.to_dict()
    print(f'Test report for {data["ProjectID"]}: {data["TestFlow"][report_number]["Test name"]} is being created.\n')

    def sub_function():

        for i in range(0, len(all_data)):
            for m in range(0, len(all_data[f"Unnamed: {i}"])):
                try:
                    if data["ProjectID"] in all_data[f"Unnamed: {i}"][m]:
                        if data["TestFlow"][report_number]["Test name"] in all_data[f"Unnamed: {i}"][m]:
                            chamber_planned = all_data[f"Unnamed: {i}"][0]
                            # print(f"Test was planned on chamber: {chamber_planned[:4]}")
                            # get chamber data for test report
                            read_data_chamber = pd.read_excel(eq_configuration, f"{chamber_planned[:4]}", header=1)
                            eq_cfg = {'Chamber': chamber_planned[:4],
                                      'Temp_system_name': read_data_chamber['Equipment Name'][0],
                                      'Temp_system_inv': read_data_chamber['Inventory/Serial No'][0],
                                      'Temp_system_calib': read_data_chamber['Calibration ID/Due date'][0],
                                      'Temp_system_qlsbz': read_data_chamber['Remarks'][0],
                                      'Ahlborn_name': read_data_chamber['Equipment Name'][1],
                                      'Ahlborn_inv': read_data_chamber['Inventory/Serial No'][1],
                                      'Ahlborn_calib': read_data_chamber['Calibration ID/Due date'][1],
                                      'Ahlborn_qlsbz': read_data_chamber['Remarks'][1],
                                      'Sensor_name': read_data_chamber['Equipment Name'][2],
                                      'Sensor_inv': read_data_chamber['Inventory/Serial No'][2],
                                      'Sensor_calib': read_data_chamber['Calibration ID/Due date'][2],
                                      'Sensor_qlsbz': read_data_chamber['Remarks'][2],
                                      }

                            return eq_cfg
                except TypeError:
                    continue

    if sub_function() is None:
        print("Name of test was not found in planning.\n")
        eq_cfg_empty = {'Chamber': "CC??",
                        'Temp_system_name': 'Name of test was not found in planning',
                        'Temp_system_inv': 'Ask Gurghean Radu to  ',
                        'Temp_system_calib': 'use the same names',
                        'Temp_system_qlsbz': 'as specified in TO',
                        'Ahlborn_name': 'N.A.',
                        'Ahlborn_inv': 'N.A.',
                        'Ahlborn_calib': 'N.A.',
                        'Ahlborn_qlsbz': 'N.A.',
                        'Sensor_name': 'N.A.',
                        'Sensor_inv': 'N.A.',
                        'Sensor_calib': 'N.A.',
                        'Sensor_qlsbz': 'N.A.',
                        }
        return eq_cfg_empty
    else:
        return sub_function()
# Search for picture at given path and return path if picture_name is at that path else return first picture.

# Format text read from QP
def standards_text_format(qp_text):
    # print(qp_text)
    if 'Failed to read standards.' in qp_text:
        return 'Failed to read standards.'
    else:
        new_text = qp_text[0].split(':',1)[1].replace("  ", "").replace(" -","-").replace("- ", "").replace(" ;",";")
        # print(new_text)
        return new_text


# Add QP snipping to the Test Report
def add_snipping(data, report_number, info_to_replace):
    info_to_replace["Snipping"] = []
    for i in range(0, len(data["TestFlow"][report_number]["QP_read_pages"])):
        # checks if value from qp_read_pages is equal with 1 and if there are more numbers. this means that we don't need first page.
        # if we let ["QP_read_pages"]) >= 1, page 1 will never be returned. To change this make > 1.
        if 1 == data["TestFlow"][report_number]["QP_read_pages"][i] and len(data["TestFlow"][report_number]["QP_read_pages"]) >= 1:
            info_to_replace["Snipping"].append({"Snipping": "Failed to read page."})
            continue
        else:
            info_to_replace["Snipping"].append({"Snipping": InlineImage(tpl,
                                                                   f'{data["TestFlow"][report_number]["Pathto04_Snipping"]}/{data["TestFlow"][report_number]["QP_read_pages"][i]}.png',
                                                                   width=Mm(150), height=Mm(193)),},)
    # print(info_to_replace['Snipping'])

def create_tracking_form(project_id, report_number):
    # Load Project ID information from database
    data = database.load_database()[project_id]
    # Get chamber data information
    chamber_data = get_chamber_data(data, planning, report_number)
    # Format text received after reading qp standard
    # print(chamber_data)
    # print(f'Values from climatic chamber are {chamber_data}')

    # Variable used to select the pictures in the folders
    path_to_picture = pathlib.Path(data['PathtoTO']).parent.parent
    path_to_test_report = f'{path_to_picture}/04_TEST REPORT/'
    path_to_picture_extended = f'{path_to_picture}/02_RAW DATA/{data["TestFlow"][report_number]["TestNo"]}'
    # print(data)

    def deviation_details_specific():
        if str(data["TestFlow"][report_number]["Test name"]) in str(data["DeviationDetails"]):
            return data["DeviationDetails"]
        else:
            # print(data["TestFlow"][report_number]["TestDeviation"])
            return data["TestFlow"][report_number]["TestDeviation"]  # 'N.A.'
    # Data which appear in the Test Report Template
    info_to_replace = {
        # --------------------------------Header------------------------------------------------------------
        'Test_method': test_method(str(data["TestFlow"][report_number]["Test name"])),
        'header': data["TestFlow"][report_number]["TestNo"],

        'SA': data["TestFlow"][report_number]["SampleAmount"],  # Sample Amount
        'SampleIdentification': data["TestFlow"][report_number]["SampleIdentification"],
        # --------------------------------Second Page---------------------------------------------------------
        'Test_name': data["TestFlow"][report_number]["Test name"],
        'DeviationDetails': deviation_details_specific(),
        'Standards':standards_text_format(data["TestFlow"][report_number]["QP_read_standards_page"]),
        # --------------------------------Third Page---------------------------------------------------------
        'Chamber': chamber_data['Chamber'],
        'Temp_system_name': chamber_data['Temp_system_name'],
        'Temp_system_inv': chamber_data['Temp_system_inv'],
        'Temp_system_calib': chamber_data['Temp_system_calib'],
        'Temp_system_qlsbz': chamber_data['Temp_system_qlsbz'],
        'Ahlborn_name': chamber_data['Ahlborn_name'],
        'Ahlborn_inv': chamber_data['Ahlborn_inv'],
        'Ahlborn_calib': chamber_data['Ahlborn_calib'],
        'Ahlborn_qlsbz': chamber_data['Ahlborn_qlsbz'],
        'Sensor_name': chamber_data['Sensor_name'],
        'Sensor_inv': chamber_data['Sensor_inv'],
        'Sensor_calib': chamber_data['Sensor_calib'],
        'Sensor_qlsbz': chamber_data['Sensor_qlsbz'],

    }
    add_snipping(data, report_number, info_to_replace)

    # print(info_to_replace)
    # Function which replace template strings with above-mentioned data
    tpl.render(info_to_replace)
    # Save and create the file in the location and with the name specified between ()
    tpl.save(f'{pathlib.Path(data["PathtoTO"]).parent}/{info_to_replace["header"]}.docx')


# create_tracking_form("R02074", "6") #used only for testing
