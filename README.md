# Attendance Program Instructions

_This program used code from [this](https://www.youtube.com/watch?v=sz25xxF_AVE) video as inspiration. Windows and Excel are required to run this program_

## Files and Folders
### Required 
1. `Attendance List` folder
2. `History` folder
3. `data.txt`
4. `set_camera.py`
5. `encode.py`
6. `detect.py`
### Created during runtime
1. `Attendance List.pkl`
2. `SOURCE.xlsx`
3. `Attendance Log.xlsx`
4. `Attendance Times.xlsx`
5. `mm-dd-yy.txt`


## Quick Start Guide for First-Time Users
### 1. Add images of faces to the `Attendance List` folder 
These images should be named in the format of `Firstname Lastname.jpg`. They will serve as the source image 
of the person during the facial recognition process. Images of non-smiling faces work best. Please only
upload one image per person.

### 2. Run `set_camera.py`
Run the `set_camera.py` file. This will display a picture from your webcam(s). Press `space` to 
cycle through your webcam(s) and note the red number on the screen. Once you know the correct webcam, press `'B'` 
on your keyboard and enter the corresponding number into the terminal window (for most users with only one webcam, 
this number will be 0). A message reading "Camera Set Successful" will be displayed in the window when complete. The
`data.txt` file containing the camera number will be automatically updated.

### 3. Run `encode.py`
Run the `encode.py` file. A message reading "Encoding Complete" will be displayed in the window 
when complete. An `Attendance List.pkl` file storing the encodings and an `SOURCE.xlsx` file will be 
automatically created.

### 4. Run `detect.py` 
Run the `detect.py` file and follow the instructions in the program. A window playing video from the webcam will be displayed. If a face is 
recognized, a green box along with the person's name will be drawn around the person's face in the video 
(unrecognized faces will have no effect ). A sign-out must occur 5 minutes or more after a sign-in for the time to be 
updated (e.g., if a person signs in at 2:25 PM, they will not be able to sign out until 2:30 PM or later).
Press and hold `'q'` to enable the command terminal. With the
command terminal enabled, you can use the following commands.

1. `'manual'` Manually input the first and last name. Manually added names will be denoted with a `*`
2. `'log'` Display name list of people that have signed in to current meeting
3. `'history'` Display sign in and sign out history of current meeting
4. `'exit'` Exit the program

A message reading "Exiting Program" will be displayed in the window when complete. An `Attendance Time.xlsx` file and
a `Attendance Log.xlsx` file will be automatically created. A `.txt` file named the current date will appear in the `History` 
folder. 

## For Future Uses
In most cases, you will **only** have to run `detect.py`. If there is anything you need to modify or update within
the system, you may need to take additional steps before running `detect.py`.

## A Few Things to Know
1. Keep all GitHub files and folders under the same directory
2. **DO NOT** alter the names of any files
3. During the `detect.py` step, the video frame rate might drop if there is a face in the frame due to facial
   recognition calculations

## Advanced Options
1. **I want to add/remove a person from the system**  
Add/remove their picture from the `Attendance List` folder and run `encode.py`.
2. **I want to change my camera**  
Run `set_camera.py` and follow the instructions for that step above.
3. **I want to change my commands**  
Open the `data.txt` file and replace the corresponding command with you custom command.
Be sure to keep spacing/formatting the same and delete trailing spaces.
