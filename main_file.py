import os
import sys
import cv2
import excel
import tkinter
from capture import capture
from tkinter import messagebox
from face_detect import FaceDetect
from students_list import StudentsList

root = tkinter.Tk()
root.withdraw()

class MainFile:
	def __init__(self, class_name):
		self.i = 0
		self.class_name = class_name
		
	def capture_and_mark(self):
		sl = StudentsList(self.class_name)
		students, roll_numbers = sl.load_pkl_file()

		FaceDetectObj = FaceDetect(self.class_name)

		Yes = True
		No = False
		Cancel = None

		i = 0
		while i <= 2:
			captured_image=None
			frame=None
			
			students_present = []
			while len(students_present) == 0:
				captured_image, frame = capture()
				students_present = FaceDetectObj.recognize(captured_image, roll_numbers)
				if students_present == "No Training Data":
					return

			try:
				name_student_present = students[roll_numbers.index(students_present[0])]
			except:
				messagebox.showerror("Error", "Recognized student not in database\nUnable to mark attendance")
				return

			response = messagebox.askyesnocancel("Confirm your identity", students_present[0]+'\n'+name_student_present)
			
			if response is Yes:
				wb = excel.attendance_workbook(self.class_name)
				excel.mark_present(wb, students_present, self.class_name)
				img_path=os.path.join(os.getcwd(), 'images', self.class_name, "s"+students_present[0][-2:],os.path.basename(captured_image))
				cv2.imwrite(img_path,frame)
				os.remove(captured_image)
				messagebox.showinfo("Attendance Confirmation", "Your attendance is marked!")
				break
			elif response is Cancel:
				break
			elif response is No:
				if i == 2:
					img_path=os.path.join(os.getcwd(), 'images',self.class_name,"unrecognized students", os.path.basename(captured_image))
					cv2.imwrite(img_path,frame)
					messagebox.showinfo("Unrecognized Student", "You were not recognized as any student of this class.\nYour attendance will be marked later if you really are")
					cv2.imwrite(img_path,frame)
				os.remove(captured_image)

			i += 1

if __name__ == '__main__':
        pass
	
