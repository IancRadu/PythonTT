from docxtpl import DocxTemplate, RichText, InlineImage
import pandas as pd
import datetime
import database
from docx.shared import Mm
# Path to the template used
tpl = DocxTemplate('./Template/TemplateRaport.docx')

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
            print(f"{report_number} not found in test tracking. Check spelling")
    except IndexError:
        print("Something went wrong with getting start and end date for tests.")
    except ValueError:
        print("Value received for start/end date is not a date")
        data_tt = {'Test_start': 'No start/end date found in Test tracking',
                   'Test_end': 'No start/end date found in Test tracking',
                   'Functional_check': read_data[read_data['Test Order/\nReport No.'] == report].values[0][16], }
        return data_tt
#
def get_chamber_data(data, planning, report_number):
    # Check equipment planning and return chamber name on which testis planned
    read_data = pd.read_excel(planning, "ENV2020", header=5)
    all_data = read_data.to_dict()
    print(data["ProjectID"])
    print(data["TestFlow"][report_number]["Test name"])
    for i in range(0, len(all_data)):
        for m in range(0, len(all_data[f"Unnamed: {i}"])):
            try:
                if data["ProjectID"] in all_data[f"Unnamed: {i}"][m]:
                    if data["TestFlow"][report_number]["Test name"] in all_data[f"Unnamed: {i}"][m]:
                        chamber_planned = all_data[f"Unnamed: {i}"][0]
                        # print(chamber_planned[:4])
                        # get chamber data for test report
                        read_data = pd.read_excel(eq_configuration, f"{chamber_planned[:4]}", header=1)
                        eq_cfg = {'Chamber': chamber_planned[:4],
                                  'Temp_system_name': read_data['Equipment Name'][0],
                                  'Temp_system_inv': read_data['Inventory/Serial No'][0],
                                  'Temp_system_calib': read_data['Calibration ID/Due date'][0],
                                  'Temp_system_qlsbz': read_data['Remarks'][0],
                                  'Ahlborn_name': read_data['Equipment Name'][1],
                                  'Ahlborn_inv': read_data['Inventory/Serial No'][1],
                                  'Ahlborn_calib': read_data['Calibration ID/Due date'][1],
                                  'Ahlborn_qlsbz': read_data['Remarks'][1],
                                  'Sensor_name': read_data['Equipment Name'][2],
                                  'Sensor_inv': read_data['Inventory/Serial No'][2],
                                  'Sensor_calib': read_data['Calibration ID/Due date'][2],
                                  'Sensor_qlsbz': read_data['Remarks'][2],
                                  }
                        return eq_cfg
            except TypeError:
                continue


def create_report(project_id, report_number):
    data = database.load_database()[project_id]
    chamber_data = get_chamber_data(data, planning, report_number)
    data_TT = add_test_start_end(data, report_number)
    print(data)
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
        'TestPlanName': data["TestPlanName"],
        'TestPlanVersionDate': data["TestPlanVersionDate"],
        'DeviationDetails': data["DeviationDetails"],
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
        'test_setup_picture': InlineImage(tpl, './test_setup_dummy.bmp', width=Mm(80), height=Mm(80)),
        # --------------------------------Fifth Page---------------------------------------------------------
        'graph_picture': InlineImage(tpl, './test_setup_dummy.bmp', width=Mm(150), height=Mm(63)),
        'details_picture': InlineImage(tpl, './test_setup_dummy.bmp', width=Mm(150), height=Mm(63)),
        # --------------------------------Last Page---------------------------------------------------------
        'LTTResponsible_Name': data['LTTResponsible']['Name'],
        'LTTResponsible_Function': data['LTTResponsible']['Function'],
        'LTTResponsible_Departament': data['LTTResponsible']['Departament'],
        'Customer_Function': data["ProjectEngineer"]["Function"],
        'Functional_check': data_TT['Functional_check']
    }

    # Function which replace template strings with above-mentioned data
    tpl.render(info_to_replace)
    tpl.save(f'{info_to_replace["header"]}.docx')


create_report("R02074", "6")
