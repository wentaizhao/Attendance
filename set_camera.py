import numpy as np
import cv2

print("----------Initializing----------")
num_cameras = 0
while cv2.VideoCapture(num_cameras, cv2.CAP_DSHOW).grab():
    num_cameras += 1
cv2.destroyAllWindows()

print("Press [space] for next image or \'b\' when ready")
camera = 0
while True:
    camera = camera % num_cameras
    cap = cv2.VideoCapture(camera, cv2.CAP_DSHOW)
    success, img = cap.read()
    height = int(cap.get(4))
    cv2.putText(img, str(camera), (30, height - 360), cv2.FONT_HERSHEY_SIMPLEX, 4, (0, 0, 255), 5, cv2.LINE_AA)
    cv2.imshow("Webcam", img)
    if cv2.waitKey(0) == ord('b'):
        break
    camera += 1


def is_integer(i):
    try:
        int(i)
        return True
    except ValueError:
        return False


number = input("Type the camera number: ")
while is_integer(number) is False:
    print("Invalid input. Try again.")
    number = input("Type the camera number: ")
file = open("data.txt", 'w')
file.write(f"Camera Number = {number}")
file.close()

cap.release()
cv2.destroyAllWindows()

print("Camera Set Successful")
input("Press Any Key to Close")  
