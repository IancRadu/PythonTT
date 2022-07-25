import pathlib

from docxtpl import DocxTemplate, RichText, InlineImage
import pandas as pd
import datetime
import database
from docx.shared import Mm
from test_method import test_method
# Path to the template used
tpl = DocxTemplate('./Template/TemplateRaport.docx')

# Assign path to general files to variable
planning = database.load_database()['doc_path']['planning']
output_location = database.load_database()['doc_path']['output_location']
eq_configuration = database.load_database()['doc_path']['eq_configuration']


# Function which search for Test start and end date
def add_test_start_end(data, report_number):
    read_data = pd.read_excel(output_location, "QL SBZ REL Tests Tracking", header=3)
    # For reading QL SBZ REL Tests Tracking
    report = data["TestFlow"][report_number]["TestNo"]
    try:
        if read_data[read_data['Test Order/\nReport No.'] == report].values[0][5] == report:

            t_start = str(read_data[read_data['Test Order/\nReport No.'] == report].values[0][10]).replace("00:00:00",
                                                                                                           "").strip()
            t_end = str(read_data[read_data['Test Order/\nReport No.'] == report].values[0][12]).replace("00:00:00",
                                                                                                         "").strip()
            data_tt = {'Test_start': datetime.datetime.strptime(t_start, "%Y-%m-%d").strftime("%d/%m/%Y"),
                       'Test_end': datetime.datetime.strptime(t_end, "%Y-%m-%d").strftime("%d/%m/%Y"),
                       'Functional_check': read_data[read_data['Test Order/\nReport No.'] == report].values[0][16], }
            return data_tt, print(data_tt['Test_start'])
        else:
            print(f"{report_number} not found in test tracking. Check spelling.\n")
    except IndexError:
        print("Something went wrong with getting start and end date for tests.\n")
    except ValueError:
        print("Value received for start/end date is not a date.\n")
        data_tt = {'Test_start': 'No start/end date found in Test tracking',
                   'Test_end': 'No start/end date found in Test tracking',
                   'Functional_check': read_data[read_data['Test Order/\nReport No.'] == report].values[0][16], }
        return data_tt


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
def get_picture(path_to_picture_extended, name, picture_name):
    data = f'{path_to_picture_extended}/{name}/'
    path = pathlib.Path(data)

    def sub_get_picture():

        for child in path.iterdir():

            if picture_name in str(child):
                return child
            elif 'png' or 'jpg' in str(child):
                print(f"No picture found with name: {picture_name}. First picture in the file was returned.\n")
                # print(str(child))
                return child

    if sub_get_picture() is None:
        print(f"No pictures found in {name}.\n")
        return './Template/test_setup_dummy.bmp'
    else:
        return sub_get_picture()

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

def create_report(project_id, report_number):
    # Load Project ID information from database
    data = database.load_database()[project_id]
    # Get chamber data information
    chamber_data = get_chamber_data(data, planning, report_number)
    # Format text received after reading qp standard

    # print(chamber_data)
    # print(f'Values from climatic chamber are {chamber_data}')
    data_TT = add_test_start_end(data, report_number)
    # Variable used to select the pictures in the folders
    path_to_picture = pathlib.Path(data['PathtoTO']).parent.parent
    path_to_test_report = f'{path_to_picture}/04_TEST REPORT/'
    path_to_picture_extended = f'{path_to_picture}/02_RAW DATA/{data["TestFlow"][report_number]["TestNo"]}'
    # print(data)

    def deviation_details_specific():
        if str(data["TestFlow"][report_number]["Test name"]) in str(data["DeviationDetails"]):
            return data["DeviationDetails"]
        else:
            return 'N.A.'
    # Data which appear in the Test Report Template
    info_to_replace = {
        # --------------------------------Header------------------------------------------------------------
        'header': data["TestFlow"][report_number]["TestNo"],
        # --------------------------------First_Page---------------------------------------------------------
        'Customer_name': data["ProjectEngineer"]["Name"],
        'Customer_phone': data["ProjectEngineer"]["Phone"],
        'Customer_dep': data["ProjectEngineer"]["Departament"],
        'EndCustomerOEM': data["EndCustomerOEM"],
        'ProjectName': data["ProjectName"],
        'Precompliance': data["TypeOfRequest"]["Pre-compliance"]["Checkbox"][0],
        'DV': data["TypeOfRequest"]["DV"]["Checkbox"][0],
        'PV': data["TypeOfRequest"]["PV"]["Checkbox"][0],
        'ExternalRequest': data["TypeOfRequest"]["ExternalRequest"]["Checkbox"][0],
        'BeforeGate80': data["TypeOfRequest"]["PV"]["BeforeGate80"][0],
        'AfterGate80': data["TypeOfRequest"]["PV"]["AfterGate80"][0],
        'WithPPPAP': data["TypeOfRequest"]["PV"]["WithPPPAP"][0],
        'WithoutPPAP': data["TypeOfRequest"]["PV"]["WithoutPPAP"][0],
        'Reason_details': data["TypeOfRequest"]["Reason_details"],
        'ProjectID': data["ProjectID"],
        'SA': data["TestFlow"][report_number]["SampleAmount"],  # Sample Amount
        'SampleIdentification': data["TestFlow"][report_number]["SampleIdentification"],
        # --------------------------------Second Page---------------------------------------------------------
        'Test_name': data["TestFlow"][report_number]["Test name"],
        'ChaperNo': data["TestFlow"][report_number]["ChaperNo"],
        'Test_method':test_method(str(data["TestFlow"][report_number]["Test name"])),
        'TestPlanName': data["TestPlanName"],
        'TestPlanVersionDate': data["TestPlanVersionDate"],
        'DeviationDetails': deviation_details_specific(),
        'Standards':standards_text_format(data["TestFlow"][report_number]["QP_read_standards_page"]),
        'Test_start': str(data_TT['Test_start']),
        'Test_end': str(data_TT['Test_end']),
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
        # --------------------------------Fourth Page---------------------------------------------------------
        'test_setup_picture': InlineImage(tpl, str(get_picture(path_to_picture_extended, '01_Pictures_Before_test',
                                                               'setup')), width=Mm(80), height=Mm(80)),
        # --------------------------------Fifth Page---------------------------------------------------------
        'graph_picture': InlineImage(tpl, './Template/test_setup_dummy.bmp', width=Mm(150),
                                     height=Mm(63)),
        'details_picture': InlineImage(tpl, str(get_picture(path_to_picture_extended, '03_Logs', 'details')),
                                       width=Mm(150), height=Mm(63)),
        # --------------------------------Last Page---------------------------------------------------------
        'LTTResponsible_Name': data['LTTResponsible']['Name'],
        'LTTResponsible_Function': data['LTTResponsible']['Function'],
        'LTTResponsible_Departament': data['LTTResponsible']['Departament'],
        'Customer_Function': data["ProjectEngineer"]["Function"],
        'Functional_check': data_TT['Functional_check']
    }
    add_snipping(data, report_number, info_to_replace)

    # print(info_to_replace)
    # Function which replace template strings with above-mentioned data
    tpl.render(info_to_replace)
    # Save and create the file in the location and with the name specified between ()
    tpl.save(f'{path_to_test_report}/01_In work/{info_to_replace["header"]}.docx')


# create_report("R02110", "2") #used only for testing
