import json


def add_data_to_form(data):
    print("S")

#Open database and search for project id
def test_tracking_form(project_id):
    with open("DataBase.json") as data_file:
        data = json.load(data_file)

    for key in data:
        if key == project_id:
            add_data_to_form(data[project_id])
        else:
            print(f"No data was found in database with {project_id}")

#Used for tests only
test_tracking_form("R03830")