import cv2
import pandas as pd
import face_recognition
import os
from openpyxl import Workbook

path = "Attendance List"
imageList = []
peopleNames = []
myList = os.listdir(path)

attendance_list = pd.DataFrame()

if not myList:
    print("\"Attendance List\" folder is empty. Please add images to continue")
    input("Press Any Key to Close") 
    raise SystemExit

print("----------Loading data----------")

for file in myList:
    curImg = cv2.imread(f"{path}/{file}")
    imageList.append(curImg)
    peopleNames.append(os.path.splitext(file)[0])

attendance_list["Names"] = peopleNames


def find_encodings(images):
    encoding_list = []
    counter = 1
    for image in images:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        encoding = face_recognition.face_encodings(image)[0]
        encoding_list.append(encoding)
        print(f"Face {counter}/{len(images)} complete")
        counter += 1
    return encoding_list


print("----------Encoding----------")

encodeListKnown = find_encodings(imageList)
attendance_list["Encodings"] = encodeListKnown
pd.to_pickle(attendance_list, "Attendance List.pkl")
print("Encoding Complete")

wb = Workbook()
sheet = wb.active
sheet.title = "SOURCE"
for i in range(len(peopleNames)):
    cell = "A" + str(i + 2)
    sheet[cell].value = peopleNames[i]
wb.save('SOURCE.xlsx')
print("Source Sheet Created")

print("Encoding Successful")
input("Press Any Key to Close")  
