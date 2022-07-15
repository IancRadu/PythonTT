import json


# Function which try to open a json file, if the file is not found, create a new file with an empty dictionary
def load_database():
    try:
        with open("DataBase.json", "r") as data_file:
            data = json.load(data_file)
            # print("DataBase loaded")
            return data
    except FileNotFoundError:
        with open("DataBase.json", "w") as data_file:
            data = {}
            json.dump(data, data_file, indent=4)
            print("New DataBase created")
            return data

#Function which updates the json file with new data
def upload_database(new_data):
    try:
        with open("DataBase.json", "r") as data_file:
            data = json.load(data_file)
            data.update(new_data)
        with open("DataBase.json", "w") as data_file:
            json.dump(data, data_file, indent=4)
    except FileNotFoundError:
        with open("DataBase.json", "w") as data_file:
            data = new_data
            json.dump(data, data_file, indent=4)
            print("New DataBase created")
            return data
#Function which check if the current project is in database, if not it will add it
def database(object):
    all_projects = load_database()
    if all_projects is None:
        all_projects = {object['ProjectID']: object}
        upload_database(all_projects)
        print(all_projects)
        return all_projects
    else:
        all_projects[f"{object['ProjectID']}"] = object
        upload_database(all_projects)
        print(all_projects)
        return all_projects
