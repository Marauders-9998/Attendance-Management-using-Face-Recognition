import pickle
import os
from os import remove
from os import getcwd
from os.path import join
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

names = []
rolls = []

class StudentsList:
    def __init__(self, class_name):
        self.class_name = class_name

    def make_pkl_file(self):
        pkl_file_path=Path(self.make_pkl_name(self.class_name))
        if pkl_file_path.exists():
            os.remove(pkl_file_path)
        wb = load_workbook(self.make_xl_name(self.class_name))
        ws = wb.active
        number_of_studs = ws['A1'].value
        for i in range(2, number_of_studs+2):
            names.append(ws['A'+str(i)].value)
            rolls.append(ws['B'+str(i)].value)
            
        with open(self.make_pkl_name(self.class_name), 'wb') as f:
            tupl = (names, rolls)
            pickle.dump(tupl, f, protocol = pickle.HIGHEST_PROTOCOL)

    def load_pkl_file(self):
        with open(self.make_pkl_name(self.class_name), 'rb') as f:
            return pickle.load(f)

    def make_xl_name(self, class_name):
        return join(getcwd(), "student's list", class_name + '.xlsx')
    
    def make_pkl_name(self, class_name):
        return join(getcwd(), "student's list", class_name + '.pkl')


if __name__ == '__main__':
	pass
