import csv
import math
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename
from sys import exit
from os.path import basename, splitext
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, colors
from openpyxl.styles.borders import Border, Side

##################
# Global variables
root = tk.Tk()
##################


def main():
    file, filename = get_file()
    wb = load_workbook(filename)
    wb1 = split_tabs(wb)
    make_plate_layout(wb1)
    make_cv_table(wb1)
    save_file(file, filename, wb1)


# Gets XLSX input file from User
def get_file():
    global root
    root.withdraw()

    success = False
    while not success:
        try:
            filename = askopenfilename(title='Choose your data files',
                                       multiple=False, filetypes=(('XLSX Files', '*.xlsx'), ('All Files', '*.*')))
            if not filename:
                exit()
            elif not filename.endswith('xlsx'):
                success = False
                messagebox.showerror(message="Invalid Filetype.",
                                     title="Failure")
            else:
                success = True
        except csv.Error as error:
            messagebox.showerror(message="Invalid Filetype.",
                                 title="Failure")

    file = open(filename)
    return file, filename


# Saves the new, more organized workbook
def save_file(file, filename, wb):
    new_name = splitext(basename(filename))[0]
    options = {}
    options['defaultextension'] = ".xlsx"
    options['filetypes'] = (('xlsx files', '*.xlsx'), ('all files', '*.*'))
    # options['initialdir'] = ""
    options['initialfile'] = new_name
    options['title'] = "Save as..."

    dest_filename = asksaveasfilename(**options)

    try:
        wb.save(filename=dest_filename)
    except PermissionError as e:
        messagebox.showerror(message="The file you are trying to overwrite is open. Close it and try again",
                             title="Failure")
        exit()

    if not dest_filename:
        exit()


# Divides the huge chunk of data into easier to read, separate sheets
def split_tabs(wb):
    ws = wb.worksheets[0]
    wb2 = Workbook()
    max_col = ws.max_column
    max_row = ws.max_row

    # Creates raw data sheet
    for i in range(1, max_row + 1):
        for j in range(1, max_col + 1):
            wb2.worksheets[0].cell(row=i, column=j).value = ws.cell(row=i, column=j).value
    wb2.worksheets[0].title = 'RAW Data'

    # Creates a sheet for every DataType
    counter = 1
    arr = []
    for i in range(1, max_row + 1):
        if ws.cell(row=i, column=1).internal_value == 'DataType:':
            wb2.create_sheet()
            data_type = str(ws.cell(row=i, column=2).internal_value)
            data_type = data_type.replace('/', ' & ')
            arr.append(i)
            wb2.worksheets[counter].title = data_type
            counter += 1

    # Fills in all but one DataType sheet with corresponding data
    for i in range(0, len(arr) - 1):
        row_num = 0
        for x in range(arr[i], arr[i + 1]):
            row_num += 1
            for y in range(1, max_col + 1):
                sheet = wb2.worksheets[i + 1]
                sheet.cell(row=row_num, column=y).value = ws.cell(row=x, column=y).value

    # Fills in the last Data type sheet with corresponding data (may be a better way to do this)
    row_num = 0
    for x in range(arr[len(arr) - 1], max_row + 1):
        row_num += 1
        for y in range(1, max_col):
            sheet = wb2.worksheets[len(arr)]
            sheet.cell(row=row_num, column=y).value = ws.cell(row=x, column=y).value

    return wb2


# Creates Plate Layout sheet
def make_plate_layout(wb):
    plate_layout = wb.create_sheet()
    plate_layout.title = 'Plate Layout'
    # Style
    font = Font(name='Verdana', size=12)
    bold_font = Font(name='Verdana', size=14, bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for i in range(1, 10):
        plate_layout.row_dimensions[i].height = 20
        for j in range(1, 14):
            cell = plate_layout.cell(row=i, column=j)
            cell.font = font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
    for i in range(2, 10):
        cell = plate_layout.cell(row=i, column=1)
        cell.font = bold_font
        cell.value = chr(i + 63)
    for j in range(2, 14):
        cell = plate_layout.cell(row=1, column=j)
        cell.font = bold_font
        cell.value = j - 1
        plate_layout.column_dimensions[chr(j + 64)].width = 20

    # Fills in Plate Layout sheet
    ws = wb.worksheets[1]
    arr = []
    for i in range(3, ws.max_row + 1):
        arr.append(ws.cell(row = i, column = 2).value)

    counter = 0
    for j in range(2, 14):
        for i in range(2, 10):
            cell = plate_layout.cell(row = i, column = j)
            cell.value = arr[counter]
            cell.alignment = Alignment(horizontal='center')
            counter += 1


# Creates %CV Table Sheet
def make_cv_table(wb):
    cv_table = wb.create_sheet()
    cv_table.title = '%CV Table'
    ws = wb.worksheets[1]

    # Finds max column
    max_col = 0
    for j in range(1, ws.max_column+1):
        if ws.cell(row=2, column=j).value == 'Total Events':
            max_col = j

    # Column width and titles
    bold_font = Font(bold=True)
    for j in range(1, max_col - 1):
        cell = cv_table.cell(row=1, column=j)
        cell.value = ws.cell(row=2, column=j+1).value
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center')
        cv_table.column_dimensions[chr(j+64)].width = 20

    # Names array
    names = []
    for i in range(3, ws.max_row + 1):
        name = ws.cell(row=i, column=2).value
        if name in names:
            pass
        else:
            names.append(ws.cell(row=i, column=2).value)

    # Fills in names
    for i in range(2, len(names) + 1):
        cv_table.cell(row=i, column=1).value = names[i-2]
        cv_table.cell(row=i, column=1).font = bold_font

    # Calculates Std Dev and Mean
    std_dev = []
    mean = []
    leave = False
    for j in range(3, max_col):
        if leave:
            break
        for i in range(3, ws.max_row, 2):
            p1 = ws.cell(row=i, column=j).value
            if p1 is None:
                leave = True
                break
            p2 = ws.cell(row=i+1, column=j).value
            avg = (p1 + p2) / 2.0
            mean.append(avg)
            p1 = abs(p1 - avg)
            p2 = abs(p2 - avg)
            p1 *= p1
            p2 *= p2
            p3 = math.sqrt(p1 + p2)
            std_dev.append(p3)

    # Calculates and fills in %CV array
    cv = []
    for i in range(0, len(std_dev)):
        x = std_dev[i]/mean[i]
        x *= 100
        x = round(x, 2)
        cv.append(x)

    # Fills in %CV into the sheet
    yellow_fill = PatternFill('solid', fgColor=colors.YELLOW)
    counter = 0
    for j in range(2, max_col + 1):
        for i in range(2, cv_table.max_row + 1):
            if counter >= len(cv):
                pass
            else:
                cell = cv_table.cell(row=i, column=j)
                cell.value = cv[counter]
                if cell.value == 0:
                    cell.fill = yellow_fill
                counter += 1

    # Border
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for j in range(1, max_col - 1):
        for i in range(1, cv_table.max_row + 1):
            cv_table.cell(row=i, column=j).border = thin_border


if __name__ == '__main__':
    exit(main())