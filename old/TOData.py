# Used to acces elements found in source word document
import docx
from simplify_docx import simplify
# access nested values using keypath using keypath_separator='.'
from benedict import benedict

class ReadTOData:

    def __init__(self, path: str):
        self.path = path
        project_data = {}

    def get_data(self):
        print(self.path)
        # read in a document
        my_doc = docx.Document(self.path)
        # coerce to JSON using the standard options
        my_doc_as_json = simplify(my_doc)
        # print(my_doc_as_json)
        # get location of all important values from my_doc_as_json using http://jsonviewer.stack.hu/
        #search for value at index, if value not found return ""
        my_doc_as_json = benedict(my_doc_as_json, keypath_separator='.')
        def get_test():
            test_list = {}

            for i in range(1, (len(my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"]))):
                test_list[i] = {
                    "Test name":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"][i]["VALUE"][0]["VALUE"][0]["VALUE"][
                            0]["VALUE"],
                    "ChaperNo":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"][i]["VALUE"][1]["VALUE"][0]["VALUE"][
                            0]["VALUE"],
                    "SampleAmount":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"][i]["VALUE"][2]["VALUE"][0]["VALUE"][
                            0]["VALUE"],
                    "SampleIdentification":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"][i]["VALUE"][3]["VALUE"][0]["VALUE"][
                            0]["VALUE"],
                    "TestDeviation":my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"][i]["VALUE"][4]["VALUE"][0]["VALUE"][0]["VALUE"],
                    "TestNo":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][0]["VALUE"][i]["VALUE"][5]["VALUE"][0]["VALUE"][
                            0]["VALUE"], }
            return test_list

        project_data = {
                "ProjectID": my_doc_as_json["VALUE"][0]["VALUE"][1]["VALUE"][0]["VALUE"],
                "ProjectName": my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][0]["VALUE"][0]["VALUE"][1]["VALUE"][0][
                    "VALUE"],
                "EndCustomerOEM": my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][0]["VALUE"][1]["VALUE"][1]["VALUE"][0][
                    "VALUE"],
                "DG Number": my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][0]["VALUE"][2]["VALUE"][1]["VALUE"][0][
                    "VALUE"],
                "ProjectEngineer": {
                    "Name":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][1]["VALUE"][1]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                    "Function":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][1]["VALUE"][0]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                    "Departament":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][1]["VALUE"][2]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                    "Phone":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][1]["VALUE"][3]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                },
                "LTTResponsible": {
                    "Name":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][2]["VALUE"][1]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                    "Function":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][2]["VALUE"][0]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                    "Departament":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][2]["VALUE"][2]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                    "Phone":
                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"][2]["VALUE"][3]["VALUE"][0]["VALUE"][
                            0][
                            "VALUE"],
                },
                "TypeOfRequest": {
                    "Pre-compliance": {
                        "Checkbox": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][2]["VALUE"][1]["VALUE"][0]["VALUE"][0]["VALUE"],
                                     my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][2]["VALUE"][1]["VALUE"][0]["VALUE"][1][
                                         "VALUE"]]},
                    "DV": {"Checkbox": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][0]["VALUE"][0],
                                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"],
                                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][3]["VALUE"][2]["VALUE"][0]["VALUE"][0]["VALUE"]]},
                    "PV": {"Checkbox": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][4]["VALUE"][1]["VALUE"][0]["VALUE"][0]["VALUE"][0],
                                        my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][4]["VALUE"][1]["VALUE"][0]["VALUE"][1]["VALUE"]],
                           "BeforeGate80": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][4]["VALUE"][2]["VALUE"][0]["VALUE"][0]["VALUE"],
                                            my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][4]["VALUE"][2]["VALUE"][0]["VALUE"][1]["VALUE"]],
                           "AfterGate80": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][5]["VALUE"][2]["VALUE"][0]["VALUE"][0]["VALUE"],
                                           my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][5]["VALUE"][2]["VALUE"][0]["VALUE"][1]["VALUE"]],
                           "WithPPPAP": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][5]["VALUE"][3]["VALUE"][0]["VALUE"][0]["VALUE"],
                                         my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][5]["VALUE"][3]["VALUE"][0]["VALUE"][1]["VALUE"]],
                           "WhithoutPPAP": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][5]["VALUE"][3]["VALUE"][1]["VALUE"][0]["VALUE"],
                                            my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][5]["VALUE"][3]["VALUE"][1]["VALUE"][1]["VALUE"]]},
                    "ExternalRequest": {
                        "Checkbox": [my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][6]["VALUE"][1]["VALUE"][0]["VALUE"][0]["VALUE"],
                                     my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][6]["VALUE"][1]["VALUE"][0]["VALUE"][1][
                                         "VALUE"]]}},
                "TestPlanName": my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][16]["VALUE"][0]["VALUE"][0]["VALUE"][0]["VALUE"].strip(
                    "TestPlanName : Test Plan (Qualification Program / Test Specification):"),
                "TestPlanVersionDate": my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][16]["VALUE"][0]["VALUE"][2]["VALUE"][0]["VALUE"].strip(
                    ": Test Plan version date: "),
                "TestFlow": get_test(),
                "DeviationDetails": my_doc_as_json["VALUE"][0]["VALUE"][3]["VALUE"][18]["VALUE"][0]["VALUE"][1]["VALUE"][0]["VALUE"].strip(
                    "Deviation details"),

            }
        for key, value in project_data.items():
                try:
                    print(f"{key} : {value}")
                except IndexError:
                    print("Got an error")
        return project_data

    def create_folder_structure(self, project_data: object) -> object:
        for key in project_data["TestFlow"]:
            print(key)
