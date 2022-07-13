import xlwings as xw  # pip install xlwings, to open an excel file and modify it
import pathlib
# Used to acces elements found in source word document
import docx
from simplify_docx import simplify
# pydash ne permite sa accesam o valoare dintr-un dict sau list priun utilizarea . (punctului) pentru cale
import pydash
# To get the path of the file
from tkinter import filedialog
# To read data from xslx file
import pandas as pd


class ReadTOData:

    def __init__(self):
        # For reading Test order information
        self.path = pathlib.PureWindowsPath(filedialog.askopenfile().name)
        # self.path ='./R03908_20220607_V02.docx' #TODO change with good one later from above
        self.test_order_file_name = pathlib.PurePath(self.path).name
        self.output_location = './CA 1014449-A13 QL SBZ REL Projects and Tests Tracking Overview.xlsx'  # TODO change with good one later
        # print(current_data)

        self.project_data = {}

    def get_data(self):
        # print(self.path)
        # read in a document
        my_doc = docx.Document(self.path)
        # coerce to JSON using the standard options
        my_doc_as_json = simplify(my_doc)

        print(my_doc_as_json)
        # get location of all important values from my_doc_as_json using http://jsonviewer.stack.hu/

        def get_test():
            test_list = {}

            for i in range(1,
                           (len(my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"]))):
                test_list[i] = {
                    "Test name":
                        pydash.get(my_doc_as_json,
                                   f'VALUE.0.VALUE.3.VALUE.18.VALUE.0.VALUE.0.VALUE.{i}.VALUE.0.VALUE.0.VALUE.0.VALUE',
                                   'N\A'),
                    "ChaperNo":
                        pydash.get(my_doc_as_json,
                                   f'VALUE.0.VALUE.3.VALUE.18.VALUE.0.VALUE.0.VALUE.{i}.VALUE.1.VALUE.0.VALUE.0.VALUE',
                                   'N\A'),
                    "SampleAmount":
                        pydash.get(my_doc_as_json,
                                   f'VALUE.0.VALUE.3.VALUE.18.VALUE.0.VALUE.0.VALUE.{i}.VALUE.2.VALUE.0.VALUE.0.VALUE',
                                   'N\A'),
                    "SampleIdentification":
                        pydash.get(my_doc_as_json,
                                   f'VALUE.0.VALUE.3.VALUE.18.VALUE.0.VALUE.0.VALUE.{i}.VALUE.3.VALUE.0.VALUE.0.VALUE',
                                   'N\A'),
                    "TestDeviation": pydash.get(my_doc_as_json,
                                                f'VALUE.0.VALUE.3.VALUE.18.VALUE.0.VALUE.0.VALUE.{i}.VALUE.4.VALUE.0.VALUE.0.VALUE',
                                                'N\A'),
                    "TestNo":
                        pydash.get(my_doc_as_json,
                                   f'VALUE.0.VALUE.3.VALUE.18.VALUE.0.VALUE.0.VALUE.{i}.VALUE.5.VALUE.0.VALUE.0.VALUE',
                                   'N\A').replace(" ", ""), }
            return test_list

        self.project_data = {
            # Search word for project ID number or get the number from path
            "ProjectID": pydash.get(my_doc_as_json, 'VALUE.0.VALUE.1.VALUE.0.VALUE',
                                    f'{self.test_order_file_name[0:6]}'),
            "ProjectName": pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.0.VALUE.0.VALUE.1.VALUE.0.VALUE', 'N\A'),
            "EndCustomerOEM": pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.0.VALUE.1.VALUE.1.VALUE.0.VALUE',
                                         'N\A'),
            "DG Number": pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.0.VALUE.2.VALUE.1.VALUE.0.VALUE', 'N\A'),
            "ProjectEngineer": {"Name": pydash.get(my_doc_as_json,
                                                   'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.1.VALUE.1.VALUE.0.VALUE.0.VALUE',
                                                   'N\A'),
                                "Function": pydash.get(my_doc_as_json,
                                                       'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.1.VALUE.0.VALUE.0.VALUE.0.VALUE',
                                                       'N\A'),
                                "Departament": pydash.get(my_doc_as_json,
                                                          'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.1.VALUE.2.VALUE.0.VALUE.0.VALUE',
                                                          'N\A'),
                                "Phone": pydash.get(my_doc_as_json,
                                                    'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.1.VALUE.3.VALUE.0.VALUE.0.VALUE',
                                                    'N\A'), },
            "LTTResponsible": {"Name": pydash.get(my_doc_as_json,
                                                  'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.2.VALUE.1.VALUE.0.VALUE.0.VALUE',
                                                  'N\A'),
                               "Function": pydash.get(my_doc_as_json,
                                                      'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.2.VALUE.0.VALUE.0.VALUE.0.VALUE',
                                                      'N\A'),
                               "Departament": pydash.get(my_doc_as_json,
                                                         'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.2.VALUE.2.VALUE.0.VALUE.0.VALUE',
                                                         'N\A'),
                               "Phone": pydash.get(my_doc_as_json,
                                                   'VALUE.0.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE.2.VALUE.3.VALUE.0.VALUE.0.VALUE',
                                                   'N\A'), },
            "TypeOfRequest": {
                "Pre-compliance": {
                    "Checkbox": [
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.2.VALUE.1.VALUE.0.VALUE.0.VALUE', 'N\A'),
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.2.VALUE.1.VALUE.0.VALUE.1.VALUE',
                                   'N\A').strip(" ")]},
                "DV": {"Checkbox": [
                    pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.3.VALUE.1.VALUE.0.VALUE.0.VALUE.0', 'N\A'),
                    pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.3.VALUE.1.VALUE.0.VALUE.1.VALUE', 'N\A').strip(
                        " "),
                    pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.3.VALUE.2.VALUE.0.VALUE.0.VALUE', 'N\A')]},
                "PV": {"Checkbox": [
                    pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.4.VALUE.1.VALUE.0.VALUE.0.VALUE.0', 'N\A'),
                    pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.4.VALUE.1.VALUE.0.VALUE.1.VALUE', 'N\A').strip(
                        " ")],
                    "BeforeGate80": [
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.4.VALUE.2.VALUE.0.VALUE.0.VALUE', 'N\A'),
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.4.VALUE.2.VALUE.0.VALUE.1.VALUE', 'N\A')],
                    "AfterGate80": [
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.5.VALUE.2.VALUE.0.VALUE.0.VALUE', 'N\A'),
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.5.VALUE.2.VALUE.0.VALUE.1.VALUE', 'N\A')],
                    "WithPPPAP": [
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.5.VALUE.3.VALUE.0.VALUE.0.VALUE', 'N\A'),
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.5.VALUE.3.VALUE.0.VALUE.1.VALUE', 'N\A')],
                    "WithoutPPAP": [
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.5.VALUE.3.VALUE.1.VALUE.0.VALUE', 'N\A'),
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.5.VALUE.3.VALUE.1.VALUE.1.VALUE', 'N\A')]},
                "ExternalRequest": {
                    "Checkbox": [
                        pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.6.VALUE.1.VALUE.0.VALUE.0.VALUE', 'N\A'),
                        pydash.get(
                            my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.6.VALUE.1.VALUE.0.VALUE.1.VALUE', 'N\A').strip(
                            " ")]},"Reason/details":pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.7.VALUE.1.VALUE.0.VALUE.0.VALUE', 'N\A').replace('Reason/details:',''),"Result":"Add_Function_To_Get_qualification_phase"},
            "TestPlanName": pydash.get(my_doc_as_json, 'VALUE.0.VALUE.3.VALUE.16.VALUE.0.VALUE.0.VALUE.0.VALUE').strip(
                "TestPlanName : Test Plan (Qualification Program / Test Specification):"),
            "TestPlanVersionDate": pydash.get(my_doc_as_json,
                                              'VALUE.0.VALUE.3.VALUE.16.VALUE.0.VALUE.2.VALUE.0.VALUE').strip(
                ": Test Plan version date: "),
            "TestFlow": get_test(),
            "DeviationDetails": pydash.get(my_doc_as_json,
                                           'VALUE.0.VALUE.3.VALUE.18.VALUE.0.VALUE.1.VALUE.0.VALUE','N/A').strip(
                "Deviation details"),
            "ProjectTrackingData": self.get_project_tracking_data(
                pydash.get(my_doc_as_json, 'VALUE.0.VALUE.1.VALUE.0.VALUE', f'{self.test_order_file_name[0:6]}')),
            "PathtoTO": str(self.path),

        }

        # for key, value in self.project_data.items():
        #     try:
        #         print(f"{key} : {value}")
        #     except IndexError:
        #         print("Got an error")
        return self.project_data

    def create_folder_structure(self, project_data: object):
        def create_folders(name):
            folder_names = ["01_Pictures_Before_test", "02_Pictures_After_test", "03_Logs", "04_Snipping"]
            # print(name)
            for i in range(0, len(folder_names)):
                pathlib.Path(f'{self.path.parent.parent}\\02_RAW DATA\\{name}\\{folder_names[i]}').mkdir(parents=True,
                                                                                                         exist_ok=True)

        for key in project_data["TestFlow"]:
            if project_data["ProjectID"] in project_data["TestFlow"][key]["TestNo"]:
                create_folders(project_data["TestFlow"][key]["TestNo"])
            else:
                create_folders(f'{project_data["ProjectID"]}{project_data["TestFlow"][key]["TestNo"]}')

    def get_project_tracking_data(self, project_id):
        # For reading Project Tracking data
        read_data = pd.read_excel(self.output_location, "QL SBZ REL Projects Tracking", header=4)
        print(project_id)
        project_tracking = {
            "BA": read_data[read_data.Column1 == project_id].values[0][6],
            "Phase": read_data[read_data.Column1 == project_id].values[0][11],
            "ValidationEngineer": read_data[read_data.Column1 == project_id].values[0][18],
            "TestEngineer": read_data[read_data.Column1 == project_id].values[0][19]
            }

        return project_tracking

    def complete_test_tracking(self, project_id):

        current_occupied_index = 4
        read_data = pd.read_excel(self.output_location, "QL SBZ REL Tests Tracking", header=3)
        # For reading QL SBZ REL Tests Tracking
        # print(f"{read_data.loc[18, 'Unique Identification No.']} didnt work")
        try:
            if read_data[read_data["Unique Identification No."] == project_id].values[0][3] == project_id:
                print(f"{project_id} is already written in test tracking")
        except IndexError:
            # print(f'{len(read_data["Unique Identification No."])}')
            def add_data(index):

                wb.sheets["QL SBZ REL Tests Tracking"][f'C{index + 5}'].value = \
                    self.project_data["ProjectTrackingData"][
                        "BA"]
                wb.sheets["QL SBZ REL Tests Tracking"][f'D{index + 5}'].value = self.project_data["ProjectID"]
                wb.sheets["QL SBZ REL Tests Tracking"][f'E{index + 5}'].value = self.project_data["ProjectName"]
                if self.project_data["ProjectID"] in self.project_data["TestFlow"][key]["TestNo"]:
                    wb.sheets["QL SBZ REL Tests Tracking"][f'F{index + 5}'].value = self.project_data["TestFlow"][key][
                        "TestNo"]
                else:
                    wb.sheets["QL SBZ REL Tests Tracking"][
                        f'F{index + 5}'].value = f'{self.project_data["ProjectID"]}{self.project_data["TestFlow"][key]["TestNo"]}'
                wb.sheets["QL SBZ REL Tests Tracking"][f'G{index + 5}'].value = \
                self.project_data["ProjectTrackingData"]["Phase"]
                wb.sheets["QL SBZ REL Tests Tracking"][f'H{index + 5}'].value = self.project_data["TestFlow"][key][
                    "Test name"]
                wb.sheets["QL SBZ REL Tests Tracking"][f'I{index + 5}'].value = \
                    self.project_data["ProjectTrackingData"]["TestEngineer"]
                wb.sheets["QL SBZ REL Tests Tracking"][f'J{index + 5}'].value = "TBD"
                wb.sheets["QL SBZ REL Tests Tracking"][f'K{index + 5}'].value = "TBD"
                # wb.sheets["QL SBZ REL Tests Tracking"][f'L{i + 5}'].value = f'=IF(ISBLANK(K357);"TBD";IF(ISTEXT(K357);"TBD";(ISOWEEKNUM(K357))))'
                # TODO = equal sign result in error, search how to send = to excel without using the sign
                wb.sheets["QL SBZ REL Tests Tracking"][f'M{index + 5}'].value = "TBD"
                wb.sheets["QL SBZ REL Tests Tracking"][f'O{index + 5}'].value = "Next"
                if self.project_data["LTTResponsible"]["Name"] != "N\A":
                    wb.sheets["QL SBZ REL Tests Tracking"][f'P{index + 5}'].value = self.project_data["LTTResponsible"][
                        "Name"]
                else:
                    wb.sheets["QL SBZ REL Tests Tracking"][f'P{index + 5}'].value = \
                        self.project_data["ProjectEngineer"]["Name"]
                xw.Range(f"{index + 6}:{index + 6}").insert("down")

            for i in range(current_occupied_index, len(read_data["Unique Identification No."])):
                # If cell from column "Uniq Ident No" is empty, return first empty row number
                if pd.isnull(read_data.loc[i, 'Unique Identification No.']):
                    # print(i)

                    wb = xw.Book(self.output_location)
                    for key in self.project_data["TestFlow"]:
                        add_data(i)
                        i += 1
                    wb.save()
                    wb.close()
                    return "Test tracking done"
                # break

