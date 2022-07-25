from docxtpl import DocxTemplate, RichText
import pandas as pd
import datetime
import database

tpl = DocxTemplate('./Template/TemplateRaport.docx')

planning = 'C:/Users/iancr/OneDrive - Continental AG/TRApp/A-13 Planning/CAP1013201-F13 Equipment Planning_20210222.xlsx'  # TODO:Add function to select path for the first time
output_location = './CA 1014449-A13 QL SBZ REL Projects and Tests Tracking Overview.xlsx'  # TODO change with good one later


# Function which search for Test start and end date
def add_test_start_end(data, report_number):
    read_data = pd.read_excel(output_location, "QL SBZ REL Tests Tracking", header=3)
    # For reading QL SBZ REL Tests Tracking
    try:
        report = data["TestFlow"][report_number]["TestNo"]
        if read_data[read_data['Test Order/\nReport No.'] == report].values[0][5] == report:
            t_start = str(read_data[read_data['Test Order/\nReport No.'] == report].values[0][10]).replace("00:00:00",
                                                                                                           "").strip()
            t_end = str(read_data[read_data['Test Order/\nReport No.'] == report].values[0][12]).replace("00:00:00",
                                                                                                         "").strip()
            data_TT = {'Test_start': datetime.datetime.strptime(t_start, "%Y-%m-%d").strftime("%d/%m/%Y"),
                       'Test_end': datetime.datetime.strptime(t_end, "%Y-%m-%d").strftime("%d/%m/%Y"),
                       'Functional_check': read_data[read_data['Test Order/\nReport No.'] == report].values[0][16], }
            return data_TT
        else:
            print(f"{report_number} not found in test tracking. Check spelling")
    except IndexError:
        print("Something went wrong with getting start and end date for tests.")

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
            except TypeError:
                continue
    #get chamber data for test report

def create_report(project_id, report_number):
    data = database.load_database()[project_id]
    get_chamber_data(data, planning, report_number)
    data_TT = add_test_start_end(data, report_number)
    print(data)
    # Data which appear in the Test Report Template
    try:
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
            # --------------------------------Last Page---------------------------------------------------------
            'LTTResponsible_Name': data['LTTResponsible']['Name'],
            'LTTResponsible_Function': data['LTTResponsible']['Function'],
            'LTTResponsible_Departament': data['LTTResponsible']['Departament'],
            'Customer_Function': data["ProjectEngineer"]["Function"],
            'Functional_check': data_TT['Functional_check']
        }
    except TypeError:
        print("")
    finally:
        # Function which replace template strings with above-mentioned data
        tpl.render(info_to_replace)
        tpl.save(f'{info_to_replace["header"]}.docx')


create_report("R02074", "6")
