import os
from os import path
import shutil
import json
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

# Open xlsx file
open_sheet = path.exists("cache/opened_sheet.json")
if open_sheet == True :
  opened_sheet_file_path = "cache/opened_sheet.json"
  json_file = open(opened_sheet_file_path)
  data = json.load(json_file)
  xlsx_sheet_check = path.exists(data ['xlsx_file_path'])
  if xlsx_sheet_check == True :
    xlsx_file_path = data ['xlsx_file_path']
  else :
    shutil.rmtree('cache', ignore_errors=True)
    xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
    cache_path = os.path.join(str(os.getcwd()), "cache")
    dictionary = {"xlsx_file_path" : xlsx_file_path}
    json_object = json.dumps(dictionary, indent = 1)
    with open("cache/opened_sheet.json", "w") as outfile:
      outfile.write(json_object)
else :
  xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
  cache_path = os.path.join(str(os.getcwd()), "cache")
  os.mkdir(cache_path)
  dictionary = {"xlsx_file_path" : xlsx_file_path}
  json_object = json.dumps(dictionary, indent = 1)
  with open("cache/opened_sheet.json", "w") as outfile:
    outfile.write(json_object)

# read imported xlsx file path using pandas
input_workbook = pd.read_excel(xlsx_file_path, sheet_name = 'Sheet1', dtype=str)
total_input_rows, total_input_cols = input_workbook.shape
print('Total Cards = ',total_input_rows)

'''
pesudo code for dynamic loading of variables
for heading in input_col:
    input_variables = input_workbook[heading].values.tolist()
    dictionary = {heading : input_variables}
    json_object = json.dumps(dictionary, indent = 1)
    with open("cache/"+ heading +".json", "w") as outfile:
      outfile.write(json_object)
    input_variables = "hemlo" + heading
    print (input_variables)
    input_variables = input_workbook[heading].values.tolist()
    print (input_variables)
'''

input_col = list(input_workbook.columns.values.tolist())

input_xlsx_col_A = input_workbook[input_col[0]].values.tolist()
input_xlsx_col_B = input_workbook[input_col[1]].values.tolist()
input_xlsx_col_C = input_workbook[input_col[2]].values.tolist()
input_xlsx_col_D = input_workbook[input_col[3]].values.tolist()
input_xlsx_col_E = input_workbook[input_col[4]].values.tolist()
input_xlsx_col_F = input_workbook[input_col[5]].values.tolist()
input_xlsx_col_G = input_workbook[input_col[6]].values.tolist()
input_xlsx_col_H = input_workbook[input_col[7]].values.tolist()
input_xlsx_col_I = input_workbook[input_col[8]].values.tolist()
input_xlsx_col_J = input_workbook[input_col[9]].values.tolist()

# get-output sheet to append output
output_sheet = path.exists("Output.xlsx")
if output_sheet == True :
  output_sheet_file_path = "Output.xlsx"
else :
  output_headers = input_col
  overall_output = Workbook()
  page = overall_output.active
  page.append(output_headers)
  overall_output.save(filename = 'Output.xlsx')
  output_sheet_file_path = "Output.xlsx"

def output_save():
  entry_list = [[input_xlsx_col_A[x], input_xlsx_col_B[x], input_xlsx_col_C[x], input_xlsx_col_D[x], input_xlsx_col_E[x], input_xlsx_col_F[x], input_xlsx_col_G[x], input_xlsx_col_H[x], input_xlsx_col_I[x], input_xlsx_col_J[x]]]
  output_wb = load_workbook(output_sheet_file_path)
  page = output_wb.active
  for info in entry_list:
      page.append(info)
  output_wb.save(filename='Output.xlsx')

for x in range(0, total_input_rows):
    output_save()