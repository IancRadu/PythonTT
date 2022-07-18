import os
import pathlib

import database
name = 1
folder_names = 7
data = database.load_database()['R02074']['PathtoTO']
print(data)
datas = pathlib.Path(data).parent

print(pathlib.Path(f"{datas.parent}\\works.bmp"))

var = os.listdir(datas)[0]
print(var)