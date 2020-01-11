from os import getcwd
from os.path import join
from calendar import monthrange
from students_list import StudentsList
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, date, timedelta
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

heading_font = Font(size = 20, bold = True)
alignment = Alignment(horizontal = 'center', vertical = 'center')
fill_date = PatternFill(fgColor = 'AFA5EF', fill_type = 'solid')
fill_roll = PatternFill(fgColor = 'A5EFAF', fill_type = 'solid')
fill_name = PatternFill(fgColor = 'E9EFA5', fill_type = 'solid')
fill_total = PatternFill(fgColor = '25F741', fill_type = 'solid')
fill_absent = PatternFill(fgColor = 'FF1010', fill_type = 'solid')
fill_present = PatternFill(fgColor = '10FF10', fill_type = 'solid')

gap = 5
date_column = 5

names = ['Marauders']
rolls = ['000001']

NUMBER_OF_STUDENTS = len(names)

def todays_date():
	return datetime.now().date()

def month():
	return todays_date().strftime('%B')

def make_file_name(class_name):
	return join(getcwd(), 'extras', class_name, class_name+ '.xlsx')

def current_month_exists(wb):
	try:
		ws = wb[month()]
		return True
	except:
		return False

def set_border(ws, cell_range):
    rows = ws[cell_range]
    side = Side(border_style='thin', color="FF000000")

    rows = list(rows)
    max_y = len(rows) - 1
    for pos_y, cells in enumerate(rows):
        max_x = len(cells) - 1
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )
            if pos_x == 0:
                border.left = side
            if pos_x == max_x:
                border.right = side
            if pos_y == 0:
                border.top = side
            if pos_y == max_y:
                border.bottom = side

            if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                cell.border = border


def createWorksheet(wb):
        #"creating register"
        global names, rolls, NUMBER_OF_STUDENTS
        manth = month()
        ws = wb.create_sheet(manth)
        ws.merge_cells('A1:AN2')
        heading = 'Attendance Record for ' + manth
        heading = heading + ' '*50 + heading + ' '*50 + heading
        ws['A1'] = heading
        ws['A1'].font = heading_font
        ws['A1'].alignment = alignment

	#ws.sheet_properties.tabColor = 'FFFF66'
        ws['A3'] = 'S.No'
        ws['B3'] = 'Enrollment'
        ws['B3'].fill = fill_roll
        ws.merge_cells('C3:D3')
        ws['C3'] = 'Name'
        ws['C3'].fill = fill_name

        cellrange = 'A'+str(gap-1)+':AN'+str(gap-1)
        set_border(ws, cellrange)
        for i in range(NUMBER_OF_STUDENTS):
                ws.merge_cells(start_row = gap+i, start_column = 3, end_row = gap+i, end_column = 4)
                ws.cell(row = gap+i, column = 1, value = str(i+1))
                ws.cell(row = gap+i, column = 2, value = rolls[i]).fill = fill_roll
                ws.cell(row = gap+i, column = 3, value = names[i]).fill = fill_name

                cellrange = 'A'+str(gap+i)+':AN'+str(gap+i)
                set_border(ws, cellrange)
        today = todays_date()
        first_day = date(today.year, today.month, 1)
        days_in_month = monthrange(today.year, today.month)[1]
        for i in range(days_in_month):
                dt = first_day + timedelta(days = i)
                ws.cell(row = 3, column = date_column+i, value = dt).fill = fill_date

        first_date_column = get_column_letter(date_column)
        last_date_column = get_column_letter(date_column+days_in_month-1)
        sum_cell_column= date_column+days_in_month+1
        ws.cell(row = 3, column = sum_cell_column, value = "Total").fill = fill_total
        for i in range(NUMBER_OF_STUDENTS):
                ws.cell(row = gap+i, column = sum_cell_column,
                        value = "=SUM("+first_date_column+str(gap+i)+":"+last_date_column+str(gap+i)+")").fill = fill_total


def mark_absent(wb, studs_absent, class_name):
	ws = wb[month()]
	col = todays_date().day + date_column - 1
	for i in range(NUMBER_OF_STUDENTS):
		if ws.cell(row = gap+i, column = 2).value in studs_absent:
			ws.cell(row = gap+i, column = col, value = 0).fill = fill_absent
	wb.save(make_file_name(class_name))


def mark_present(wb, studs_present, class_name):
	ws = wb[month()]
	col = todays_date().day + date_column - 1
	for i in range(NUMBER_OF_STUDENTS):
		if ws.cell(row = gap+i, column = 2).value in studs_present:
			ws.cell(row = gap+i, column = col, value = 1).fill = fill_present
	wb.save(make_file_name(class_name))

def attendance_workbook(class_name):
    sl = StudentsList(class_name)
    tupl = sl.load_pkl_file()
    global names, rolls, NUMBER_OF_STUDENTS
    names = tupl[0]
    rolls = tupl[1]
    NUMBER_OF_STUDENTS = len(names)
    #Current length of names is len(names)
    filename = make_file_name(class_name)
    try:
        wb = load_workbook(filename)
    except:
        wb = Workbook()
    wb.guess_types = True
    if not current_month_exists(wb):
        createWorksheet(wb)
    wb.save(filename)
    return wb


if __name__ == '__main__':
        pass
