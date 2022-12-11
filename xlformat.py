import openpyxl
import tkinter as tk
from tkinter import filedialog
import os

import tkinter as tk
from tkinter import filedialog
from openpyxl import workbook
from openpyxl import worksheet

root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory(title="Select Folder To Store")
path_xl = filedialog.askopenfilename(title="Select Employee Pay Schedule File") #Download the current pay schedule file and upload here. 
# path_csv = filedialog.askopenfilename(title="Select Employee Payroll Export") #Download the payroll export and upload here. 
os.chdir(file_path)

wb = load_workbook(filename = path_xl)
ws = wb.active
    # worksheet = writer.sheets['Hours']
    
header = workbook.add_format({'bold': True}) # Adding bold for the column headers
col_a = workbook.add_format({'align':'left'})
number_format = workbook.add_format({'num_format': '#,##0.00'})

worksheet.set_column(1, 2, 20, number_format)
worksheet.set_column('A:A', 30, col_a)
worksheet.set_column('C:Z', 12, None)
writer.save()