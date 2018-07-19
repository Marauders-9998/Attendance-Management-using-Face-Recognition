import excel
import os
import cv2
import sys
import excel
import shlex
import pyautogui
import subprocess
from capture import capture
from capture import detect_faces
from datetime import datetime
from calendar import monthrange
from openpyxl import Workbook, load_workbook
from face_train import FaceTrain
from students_list import StudentsList
import tkinter as tk
from tkinter import *
from main_file import MainFile
from tkinter import font as tkfont
import face_train

class_codes = ['Marauders']
manager_id = 'ADMIN'
manager_pass = 'ubuntu'


current_class = 'Marauders'
if current_class is not 'Marauders':
	current_class_obj = MainFile(current_class)
	FaceTrainObj = FaceTrain(current_class)
else:
	current_class_obj = None
	FaceTrainObj = None


class SampleApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold", slant="italic")
        x, y = pyautogui.size()
        geo = str(int(0.35*x))+'x'+str(int(0.85*y))
        self.geometry(geo)
        self.resizable(False, False)
        self.title('Attendance Management App')
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.createInitialDirectories()

        self.frames = {}
        for F in (StartPage, StudentPanelPage, ManagerPanelPage, CreateNewBatchPage, AddStudentPage):
            page_name = F.__name__
            
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order
            # will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")

    def createInitialDirectories(self):
    	
    	temp_dir = os.path.join(os.getcwd(), 'images', '.temp')
    	temp_dir_exists = os.path.isdir(temp_dir)

    	extras_dir = os.path.join(os.getcwd(), 'extras')
    	extras_dir_exists = os.path.isdir(extras_dir)

    	studs_list_dir = os.path.join(os.getcwd(), "student's list")
    	studs_list_dir_exists = os.path.isdir(studs_list_dir)

    	if not extras_dir_exists:
    		os.makedirs(extras_dir)
    	if not temp_dir_exists:
    		os.makedirs(temp_dir)
    	if not studs_list_dir_exists:
    		os.makedirs(studs_list_dir)

    	classes_file = os.path.join(os.getcwd(), 'extras', 'classes.xlsx')
    	try:
    		global class_codes, current_class
    		class_codes = set(['Marauders'])
    		wb = load_workbook(classes_file)
    		ws = wb.active
    		number_of_classes = ws['A1'].value
    		if number_of_classes != 0:
    			current_class=ws['A2'].value

    		for i in range(2,number_of_classes+2):
    			class_codes.add(ws['A'+str(i)].value)
    	except:
    		wb = Workbook()
    		wb.guess_types = True
    		ws = wb.active
    		ws['A1'] = 0
    		wb.save(classes_file)

    def show_frame(self, page_name):
        #Show a frame for the given page name
        frame = self.frames[page_name]
        frame.tkraise()

    def on_closing(self):
    	self.destroy()
    	self.quit()

    def validate(self, page_name,c_code,id,pwd):
    	global class_codes
    	if id==manager_id and pwd==manager_pass and c_code in class_codes:
    		frame = self.frames[page_name]
    		global current_class
    		current_class = c_code
    		print("current class",current_class)
    		if current_class != 'Marauders':
    			#frame.lb_class_code.config(text="Batch Code : {}".format(current_class))
    			#frame.lb_class_code['text']="Batch Code : {}".format(current_class)
    			StudentsList(current_class).make_pkl_file()
    			if page_name == "ManagerPanelPage":
    				global FaceTrainObj
    				FaceTrainObj = FaceTrain(current_class)
    			elif page_name == "StudentPanelPage":
    				global current_class_obj
    				current_class_obj = MainFile(current_class)
    		elif current_class == 'Marauders' and page_name == 'StudentPanelPage':
    			messagebox.showerror('Error','Only Manager Panel is valid for this class-code')
    			return
    		else:
    			frame = self.frames['CreateNewBatchPage']
    		frame.tkraise()
    	else:
    		tk.messagebox.showerror('Error','Incorrect Credentials !!')


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        
        self.label = tk.Label(self, text="Attendance Management", font=controller.title_font)
        self.label.pack(side="top", fill="x", pady=10)
        self.lb_class=tk.Label(self,text="CLASS-CODE : ",width=15)
        self.lb_class.pack()
        self.tv_class=tk.Entry(self,width=15)
        self.tv_class.focus()
        self.tv_class.pack()
        self.lb_username=tk.Label(self,text="USERNAME : ",width=15)
        self.lb_username.pack()
        self.tv_username=tk.Entry(self,width=15)
        self.tv_username.pack()
        self.lb_pass=tk.Label(self,text="PASSWORD : ",width=15)
        self.lb_pass.pack()
        self.tv_pass=tk.Entry(self,show="*",width=15)
        self.tv_pass.pack()

        button1 = tk.Button(self, text="Go to Student Portal", width=30,
                            command=lambda: self.doWork("StudentPanelPage",self.tv_class.get(),self.tv_username.get(),self.tv_pass.get()))
        button2 = tk.Button(self, text="Go to Manager Portal", width=30,
                            command=lambda: self.doWork("ManagerPanelPage",self.tv_class.get(),self.tv_username.get(),self.tv_pass.get()))
        button1.pack()
        button2.pack()
        bt_exit = tk.Button(self, text="Exit",width=20,
                           command=self.exit)
        bt_exit.pack()

    def doWork(self, page_name,c_code,id,pwd):
    	self.tv_class.delete(0, END)
    	self.tv_username.delete(0, END)
    	self.tv_pass.delete(0, END)
    	self.controller.validate(page_name, c_code, id, pwd)

    def exit(self):
        self.controller.destroy()
        self.controller.quit()
            


class StudentPanelPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="Student Portal", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)
        
        self.lb_class_code=tk.Label(self,width=30).pack()
        bt_mark = tk.Button(self, text="Mark my attendance",width=30,
                           command = self.doWork)
        bt_mark.pack()

        bt_back = tk.Button(self, text="Back",width=10,
                           command=lambda: controller.show_frame("StartPage"))
        bt_back.pack()

        bt_exit = tk.Button(self, text="Exit",width=10,
                           command=self.exit)
        bt_exit.pack()

    def doWork(self):
        global current_class_obj
        current_class_obj.capture_and_mark()

    def exit(self):
        self.controller.destroy()
        self.controller.quit()


class ManagerPanelPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="Manager Portal", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)
        global current_class
        self.lb_class_code=tk.Label(self, width=30).pack()
        
        bt_train = tk.Button(self, text="Train the Recogniser",width=30,command = self.doWork)
        bt_train.pack()
        bt_addstud = tk.Button(self, text="Add a student to this class",width=30,command = self.addStudent)
        bt_addstud.pack()
        bt_view_register = tk.Button(self, text="View Attendance Register",width=30,command = self.viewRegister)
        bt_view_register.pack()
      
        bt_back = tk.Button(self, text="Back",width=30,
                           command=lambda: controller.show_frame("StartPage"))
        bt_back.pack()
        bt_exit = tk.Button(self, text="Exit",width=30,
                           command=self.exit)
        bt_exit.pack()

    def viewRegister(self):
    	global current_class
    	atten_register = os.path.join(os.getcwd(), 'extras', current_class, current_class+'.xlsx')
    	try:
    		opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
    		subprocess.call([opener, atten_register])
    	except:
    		os.startfile(atten_register)


    def addStudent(self):
    	global current_class
    	students_list_file = os.path.join(os.getcwd(), "student's list", current_class+'.xlsx')
    	self.controller.show_frame("AddStudentPage")
    
    def doWork(self):
        global FaceTrainObj, current_class
        if current_class == 'Marauders':
        	wb = load_workbook(os.path.join(os.getcwd(),'extras','classes.xlsx'))
        	ws = wb.active
        	number_of_classes = ws['A1'].value
        	if number_of_classes==0:
        		messagebox.showerror('Error','No class in the database yet')
        		return
        	else:
        		messagebox.showerror('Error','Please login again with a valid class-code')
        		current_class=ws['A2'].value
        		return

        FaceTrainObj.main()

    def exit(self):
        self.controller.destroy()
        self.controller.quit()


class CreateNewBatchPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        label = tk.Label(self, text="Create a new Batch", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)

        lb_class=tk.Label(self,text="CLASS-CODE : ",width=15)
        lb_class.pack()
        self.tv_class=tk.Entry(self,width=15)
        self.tv_class.focus()
        self.tv_class.pack()
        lb_number=tk.Label(self,text="Number of students : ")
        lb_number.pack()

        vcmd = (self.controller.register(self.validate_number_field), '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        self.tv_number=tk.Entry(self, width=15, validate = 'key', validatecommand = vcmd)
        self.tv_number.pack()

        bt_new_batch = tk.Button(self, text="Add Batch",width=30, command = self.create_new_batch)
        bt_new_batch.pack()
       
        bt_back = tk.Button(self, text="Back",width=30,
                           command=lambda: controller.show_frame("StartPage"))
        bt_back.pack()


    def validate_number_field(self, action, index, value_if_allowed, prior_value, text, validation_type, trigger_type, widget_name):
    	if text in '0123456789':
    		try:
    			int(value_if_allowed)
    			return True
    		except:
    			return False
    	else:
    		return False

    def create_new_batch(self):
    	global class_codes
    	class_name = self.tv_class.get()
    	number_of_studs = int(self.tv_number.get())

    	batchexists = os.path.exists(os.path.join(os.getcwd(), 'extras', class_name))
    	if batchexists:
    		tk.messagebox.showerror('Error', 'This batch name already exists')
    		return

    	if number_of_studs < 1 or number_of_studs > 99:
    		tk.messagebox.showerror('Error', "Number of students not in allowed range!")
    		return    		

    	images_dir = os.path.join(os.getcwd(), 'images', class_name)
    	unrecog_studs_dir = os.path.join(os.getcwd(), 'images', class_name,'unrecognized students')
    	os.makedirs(unrecog_studs_dir)

    	for i in range(number_of_studs):
    		os.makedirs(os.path.join(images_dir, 's'+str(i).zfill(2)))

    	studs_list_file = os.path.join(os.getcwd(), "student's list", class_name+'.xlsx')
    	wb = Workbook()
    	wb.guess_types = True
    	ws = wb.active
    	ws['A1'] = 0
    	ws['B1'] = number_of_studs
    	wb.save(studs_list_file)

    	classes_file = os.path.join(os.getcwd(), 'extras', 'classes.xlsx')
    	wb = load_workbook(classes_file)
    	ws = wb.active
    	row = ws['A1'].value
    	ws['A1']=row+1
    	ws['A'+str(row+2)] = class_name
    	wb.save(classes_file)

    	sl = StudentsList(class_name)
    	sl.make_pkl_file()

    	atten_reg_dir = os.path.join(os.getcwd(), 'extras', class_name)
    	os.makedirs(atten_reg_dir)
    	wb = excel.attendance_workbook(class_name)

    	tk.messagebox.showinfo("Batch Successfully Created", "All the necessary Directories and Files created \nfor Batch-Code {}".format(class_name))
    	class_codes.add(class_name)
    	self.tv_class.delete(0, END)
    	self.tv_number.delete(0, END)
    	self.tv_number.insert(0, "")
    	self.controller.show_frame('StartPage')


class AddStudentPage(tk.Frame):
	"""docstring for AddStudentPage"""
	def __init__(self, parent, controller):
		tk.Frame.__init__(self, parent)
		self.controller = controller

		label = tk.Label(self, text="Add a new Student", font=controller.title_font)
		label.pack(side="top", fill="x", pady=10)

		self.lb_class_code=tk.Label(self,width=30).pack()

		lb_name=tk.Label(self,text="Name of student")
		lb_name.pack()
		self.tv_name=tk.Entry(self,width=30)
		self.tv_name.focus()
		self.tv_name.pack()

		bt_addstud = tk.Button(self, text="Add Student",width=30,command = self.doWork)
		bt_addstud.pack()

		bt_back = tk.Button(self, text="Back",width=30,
                           command=lambda: controller.show_frame("ManagerPanelPage"))
		bt_back.pack()
		bt_exit = tk.Button(self, text="Exit",width=30,
                           command=self.exit)
		bt_exit.pack()

	def doWork(self):
		global current_class
		
		students_list_file = os.path.join(os.getcwd(), "student's list", current_class+'.xlsx')
		wb = load_workbook(students_list_file)
		ws = wb.active
		total_studs=ws['B1'].value
		currently_studs = ws['A1'].value
		if currently_studs==total_studs:
			messagebox.showinfo("Batch Full", "No more accomodation for any new student")
			return
		ws['A1'] = currently_studs + 1
		row_here = currently_studs + 2
		student_name = self.tv_name.get()
		ws['A'+str(row_here)] = student_name
		roll_no = currently_studs
		roll_no = str(roll_no).zfill(2)
		ws['B'+str(row_here)] = current_class+roll_no
		wb.save(students_list_file)

		sl = StudentsList(current_class)
		sl.make_pkl_file()

		atten_reg = os.path.join(os.getcwd(), 'extras', current_class, current_class+'.xlsx')
		wb = load_workbook(atten_reg)
		today = datetime.now().date()
		month = today.strftime('%B')
		try: #If current month exists
			ws = wb[month]
			row_here = excel.gap + int(roll_no)
			ws.merge_cells(start_row = row_here, start_column = 3, end_row = row_here, end_column = 4)
			ws.cell(row = row_here, column = 1, value = str(int(roll_no)+1))
			fill_roll = excel.PatternFill(fgColor = 'A5EFAF', fill_type = 'solid')
			fill_name = excel.PatternFill(fgColor = 'E9EFA5', fill_type = 'solid')
			ws.cell(row = row_here, column = 2, value = current_class + roll_no).fill = fill_roll
			ws.cell(row = row_here, column = 3, value = student_name).fill = fill_name
			cellrange = 'A'+str(row_here)+':AN'+str(row_here)
			excel.set_border(ws, cellrange)
			first_date_column = excel.get_column_letter(excel.date_column)
			days_in_month = monthrange(today.year, today.month)[1]
			last_date_column = excel.get_column_letter(excel.date_column+days_in_month-1)
			sum_cell_column= excel.date_column+days_in_month+1
			fill_total = excel.PatternFill(fgColor = '25F741', fill_type = 'solid')
			ws.cell(row = row_here, column = sum_cell_column,
				value = "=SUM("+first_date_column+str(row_here)+":"+last_date_column+str(row_here)+")").fill = fill_total
			wb.save(atten_reg)
		except: #If current month does not exists
			wb = excel.attendance_workbook(current_class)

		image_dir = os.path.join(os.getcwd(), 'images', current_class, 's'+roll_no)
		messagebox.showinfo('Student Added', student_name+" admitted!\nClick OK to proceed for capturing training images. Make sure that your surroundings are well lit")
		i=5
		face_detected_right=False
		xml_file = os.path.join(os.getcwd(),'haarcascade_frontalface_default.xml')
		while i>0:
			while face_detected_right is False:

				img_path, frame = capture()
				face_detected_right=detect_faces(xml_file,frame)
			img_path=os.path.join(os.getcwd(), 'images',current_class,"s"+str(roll_no).zfill(2), os.path.basename(img_path))
			cv2.imwrite(img_path,frame)
			cv2.destroyAllWindows()
			face_detected_right = False
			i=i-1
		if i==0	:
			messagebox.showinfo("Message","Training images added successfully!")
		messagebox.showinfo('Enrollment Number', 'Roll Number of student: '+current_class+roll_no)
		self.tv_name.delete(0, END)

	def exit(self):
		self.controller.destroy()
		self.controller.quit()


if __name__ == "__main__":
    app = SampleApp()
    app.mainloop()
