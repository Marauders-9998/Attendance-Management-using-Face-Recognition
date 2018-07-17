
# [Attendance-Management-using-Face-Recognition](https://github.com/Marauders-9998/Attendance-Management-using-Face-Recognition)

## Objective
This desktop application aims to simplify the process of Attendance Management of different classes and batches using Face Recognition and user-friendly GUI. The management of data and marking of attendance is carried out in Excel files.

## About the Project
This python based app uses OpenCV libraries: face detection using Haar feature-based cascade classifier and face recognition using Local Binary Patterns Histograms (LBPH) and openpyxl library to manage excel sheet through python scripts. GUI features are implemented using tkinter library.

## About the app
The app has two main Panels: Students' and Manager's
![image description](https://lh3.googleusercontent.com/k8wLf41-IzIm5AiUo8G44IrBsLKKKtzKUqnoNiOGEo7KHox6ky9YpiaZbBplw0lLOYrHBiZLzMm8 "Welcome Page")
<br/>
Students' panel has one main feature in marking the attendance
![image description](https://lh3.googleusercontent.com/Sq2hEKvJoPdtKrsYNakuKoBOF10utSc6nLyiKfQnVCy9DWG511sLHcIAY9MjV-WVP4hP3Sz1cOTG "Student Panel Page")
<br/>
Manager's panel has various features in adding a new class, adding a student to logged-in class and training the recognizer
![image description](https://lh3.googleusercontent.com/BQoRZ4yOKuiclxaVxreUeHHcSkekon_klQFP7HFB0BRFSSgMSdr_uV9jPcpCfa0QhnYzzpnn-ukh "Create a New Batch")
![image description](https://lh3.googleusercontent.com/jbRvtzBtTERP3MTpYrekBSjHmib_YlhgH1sSCxUXVtBVB2rojEkMUHgs4r9DtFV6y5EKA3Z6T8OC "Manager Panel Page")
<br/>
Captured photo after marking attendance automatically gets added to training data images. If the app somehow fails to recognize a student, the Manager can manually mark a student's attendance in attendance register by looking at the images in unrecognized students directory.
<br/>
**Directory Structure:**
<br/>
./extras - contains the attendance register of different classes and the file containing names of all classes
<br/>
./images - contains the training data images and unrecognized student's directory
<br/>
./student's list - contains the files containing the information of all students of a class

## Setup for the app
Python3 should be installed on the system.
Windows users can download from [here](https://www.python.org/downloads/), Linux users will most probably have python3 already installed or it can be done manually using `sudo apt-get install python3.6`
>Windows users can add python and pip to PATH (Linux users can use `sudo apt-get install pip3`)

**Install Modules**
 1. `pip install openpyxl`
 2. `pip install pillow`
 3. `pip install pyautogui`
 4. `pip install OpenCV-contrib-python`

##  How to run the app
 - run the Attendance_app.py to launch the software
 - use 'Marauders' as class-code to add classes to the software
 - browse the Manager and Student's Panel using class code of the batch you want to enter with
 - the username is 'ADMIN' and password 'ubuntu'
 - use `SPACE` to click images wherever needed
 - you can add more images to the training data manually
