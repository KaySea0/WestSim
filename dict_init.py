import openpyxl
import json
import tkinter as tk
from tkinter import messagebox

def cage_dict_init(file):
	cage_wb = openpyxl.load_workbook(file)
	cage_ws = cage_wb.active
	cage_max_row = cage_ws.max_row

	cage_dict = {}
	po_dict = {}

	for i in range (1, cage_max_row+1):
		if cage_ws['D' + str(i)].value and cage_ws['B' + str(i)].value:
			dict_entry = {'email': cage_ws['D'+str(i)].value}
			if cage_ws['A' + str(i)].value is None:
				dict_entry['options'] = ""
			else:
				dict_entry['options'] = cage_ws['A' + str(i)].value
				
			cage_dict[cage_ws['B' + str(i)].value] = dict_entry
			
	json_temp = json.dumps(cage_dict)
	f = open("cage_dict.json","w")
	f.write(json_temp)
	f.close()
	
	tk.messagebox.showinfo("Cagecode Confirmation", "Cagecode reference list has been created/updated!")