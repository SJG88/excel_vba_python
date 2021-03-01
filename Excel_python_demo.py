# -*- coding: utf-8 -*-
"""
This is a script to demo how to open up a macro enabled excel file, write a pandas dataframe to it
and save it as a new file name.

Created on Mon Mar  1 17:47:41 2021

@author: Shane Gore
"""
import os 
import xlwings as xw
import pandas as pd

os.chdir(r"C:\Users\Shane Gore\Desktop\Roisin")

wb = xw.Book("CAO_template.xlsm")
worksheet = wb.sheets['EPOS_Closing_Stock_Detailed']

'Create dataframe'
cars = {'Brand': ['Honda Civic','Toyota Corolla','Ford Focus','Audi A4'],
        'Price': [22000,25000,27000,35000]
        }

cars_df = pd.DataFrame(cars, columns = ['Brand', 'Price'])


'Write a dataframe to excel'
worksheet.range('A1').value = cars_df

'Create a datafame from and excel sheet'
excel_df = worksheet.range('A1').options(pd.DataFrame, expand='table').value


'Save the excel as a new workbook'
newfilename = ('Test4.xlsm')
wb.save(newfilename)

'close the workbook'
wb.close()


