import json

with open("corrections.json","r") as f:
    correctiondata = json.load(f)

with open("studentdata.json","r") as f:
    studentdata = json.load(f)

for course in correctiondata:
    for idx in correctiondata[course]:
        if idx not in studentdata and len(idx) != 6:
            print(f"Index {idx} in course {course} not found in student data.")