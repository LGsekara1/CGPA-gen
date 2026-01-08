from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent



srcdataPath = BASE_DIR / "data"/"student_data.txt"
src2dataPath = BASE_DIR/"data"/"bme_data.txt"
targetdataPath = BASE_DIR/"data"/"processed_data.txt"

DATA={} #TO store data after extracting from .txt files

with open(srcdataPath,"r") as f:
    student_data = f.readlines()
    #print(student_data)
    for student in student_data:
        idx,name = student.strip().split("\t")
        DATA[idx] = f"{name}"
        

    #print(student_data)
    #print(DATA)


BME_DATA={} #TO store data after extracting from .txt files

with open(src2dataPath,"r") as f:
    bme_data = f.readlines()
   #print(bme_data)
    for record in bme_data:
        #print(record.strip().split(" "))
        idx, sname, iname = record.strip().split(" ")
        BME_DATA[idx] = f"{sname} {iname}"


print(BME_DATA.keys())
# print("-----")
# print(DATA.keys())
#print(DATA)


PROCESSED_DATA={} #TO store final processed data

for raw_idx in DATA.keys():
    #print(raw_idx)
    idx = raw_idx[:-1]


    if raw_idx in BME_DATA.keys():
        spec = "BME"
        print("BME found:", idx )
    else:
        spec = "ENTC"

    PROCESSED_DATA[raw_idx] ={
        "raw_idx": raw_idx,
        "idx":idx,
        "name":DATA[raw_idx],
        "spec":spec
    }

import json
#print(json.dumps(PROCESSED_DATA, indent=4))
with open(BASE_DIR/"data"/"student_details.json","w") as f:
    json.dump(PROCESSED_DATA,f, indent=4)