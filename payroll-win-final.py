import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import os

import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory(title="Select Folder To Store")
path_xl = filedialog.askopenfilename(title="Select Employee Pay Schedule File") #Downlaod the current pay scheudle file and upload here. 
path_csv = filedialog.askopenfilename(title="Select Employee Payroll Export") #Download the payroll export and upload here. 

os.chdir(file_path)

#read Aspire payroll export
df = pd.read_csv(path_csv) 
df2 = pd.read_excel(path_xl, 'Sheet0') # Read pay schedule export into Dataframe

# print(df.head)
#print(df2.head)


# Counts the number of unique payroll agencies from the pay schedule export (not really needed)
agency = df2["Pay Schedule Name"].unique() 
agency_count = len(agency)
#print(agency_count)

# Merges the pay schedule names with the employee so each employee's hours has a pay schedule attached to it. 
df3 = df.merge(df2,how='left', left_on='Employee Name', right_on='Full Name')
#print(df3.head(25))

# Counts the number of pay schedules from the payroll export itself. Counts them and makes a list to use later for splitting everything up. 
agency = df3["Pay Schedule Name"].unique()
agency_count = len(agency)
#print(agency)
#print(agency_count)

# Pivots all of the hours into a summary
pivot = pd.pivot_table(data=df3, index=['Pay Schedule Name', 'Employee Name','E Base Rate','CC1'], values='E Hours', columns='E/D/T Code',  aggfunc=sum, 
    fill_value=0.0, margins=True, margins_name='Total Hours' )

print(pivot)

print(pivot.head(20))

# Call the Excel writer from openpxyl and create an .xls for each pay schedule
# writer = pd.ExcelWriter('payroll.xlsx')
# pivot.to_excel(writer) #, sheet_name="Summary"

#for paysched in pivot.index.get_level_values(0).unique():
#    file = (paysched,".xlsx") 
#    filename = ''.join(file)
#    writer = pd.ExcelWriter(filename)
#    temp_df = pivot.xs(paysched, level=0)
#    temp_df.to_excel(writer,paysched)
#    writer.save()

for paysched in pivot.index.get_level_values(0).unique():
    file = (paysched,".xlsx") 
    filename = ''.join(file)
    writer = pd.ExcelWriter(filename, 
        engine='xlsxwriter',
        engine_kwargs={'options':{'strings_to_numbers': True}})
    temp_df = pivot.xs(paysched, level=0)
    temp_df.to_excel(writer,sheet_name='Hours')
    workbook = writer.book
    worksheet = writer.sheets['Hours']
    header = workbook.add_format({'bold': True}) # Adding bold for the column headers
    col_a = workbook.add_format({'align':'left'})
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    worksheet.set_column(1, 2, 20, number_format)
    worksheet.set_column('A:A', 30, col_a)
    worksheet.set_column('C:Z', 12, None)
    writer.save()


