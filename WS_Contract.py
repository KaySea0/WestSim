import tkinter as tk
import json
import openpyxl
from pathlib import Path

# if no preservation method is given / is special mil spec, put "n/a" for method in display

class WS_Contract(object):

	def __init__(self):
		self.main_wb = None
		self.wip_wb = None
		self.dict = None
		
	def process_addition(self, var_list, window):
		main_ws = self.main_wb['DLAORDERS']
		next_row = main_ws.max_row+1
		
		main_ws['G'+str(next_row)] = var_list[0].get()
		main_ws['B'+str(next_row)] = var_list[1].get()
		main_ws['E'+str(next_row)] = var_list[2].get()
		main_ws['K'+str(next_row)] = var_list[3].get()
		main_ws['D'+str(next_row)] = var_list[4].get()
		main_ws['F'+str(next_row)] = var_list[6].get()
		main_ws['H'+str(next_row)] = var_list[9].get()
		
		self.main_wb.save(self.dict['main'])
		
		wip_ws = self.wip_wb.active
		next_row = wip_ws.max_row+1
		
		wip_ws['A'+str(next_row)] = var_list[0].get()
		wip_ws['B'+str(next_row)] = var_list[1].get()
		wip_ws['C'+str(next_row)] = var_list[2].get()
		wip_ws['D'+str(next_row)] = var_list[3].get()
		wip_ws['E'+str(next_row)] = var_list[4].get()
		wip_ws['F'+str(next_row)] = var_list[5].get()
		wip_ws['G'+str(next_row)] = var_list[6].get()
		wip_ws['H'+str(next_row)] = var_list[7].get()
		wip_ws['I'+str(next_row)] = var_list[8].get()
		wip_ws['J'+str(next_row)] = var_list[9].get()
		
		self.wip_wb.save(self.dict['wip'])
		
		window.destroy()
		
	def add_contract(self):
		t = tk.Toplevel()
		t.geometry('300x500')
		t.title("Add New Contract")
		
		varList = []
		for i in range(0,10):
			temp = tk.StringVar()
			varList.append(temp)
		
		date_label = tk.Label(t, text="Data Awarded:")
		date_label.grid(row=0, column=0, sticky="e", padx=10, pady=10)
		
		date_entry = tk.Entry(t, width=12, textvariable=varList[0])
		date_entry.grid(row=0, column=1, sticky="w", padx=10, pady=10)
		
		cnumber_label = tk.Label(t, text="Contract #:")
		cnumber_label.grid(row=1, column=0, sticky="e", padx=10, pady=10)
		
		cnumber_entry = tk.Entry(t, width=18, textvariable=varList[1])
		cnumber_entry.grid(row=1, column=1, sticky="w", padx=10, pady=10)
		
		qty_label = tk.Label(t, text="Quantity:")
		qty_label.grid(row=2, column=0, sticky="e", padx=10, pady=10)
		
		qty_entry = tk.Entry(t, width=5, textvariable=varList[2])
		qty_entry.grid(row=2, column=1, sticky="w", padx=10, pady=10)
		
		ctotal_label = tk.Label(t, text="Contract Total:")
		ctotal_label.grid(row=3, column=0, sticky="e", padx=10, pady=10)
		
		ctotal_entry = tk.Entry(t, width=11, textvariable=varList[3])
		ctotal_entry.grid(row=3, column=1, sticky="w", padx=10, pady=10)
		
		nsn_label = tk.Label(t, text="NSN:")
		nsn_label.grid(row=4, column=0, sticky="e", padx=10, pady=10)
		
		nsn_entry = tk.Entry(t, width=15, textvariable=varList[4])
		nsn_entry.grid(row=4, column=1, sticky="w", padx=10, pady=10)
		
		partname_label = tk.Label(t, text="Part Name:")
		partname_label.grid(row=5, column=0, sticky="e", padx=10, pady=10)
		
		partname_entry = tk.Entry(t, width=30, textvariable=varList[5])
		partname_entry.grid(row=5, column=1, sticky="w", padx=10, pady=10)
		
		vendor_label = tk.Label(t, text="Vendor Name:")
		vendor_label.grid(row=6, column=0, sticky="e", padx=10, pady=10)
		
		vendor_entry = tk.Entry(t, width=25, textvariable=varList[6])
		vendor_entry.grid(row=6, column=1, sticky="w", padx=10, pady=10)
		
		pn_label = tk.Label(t, text="Part #:")
		pn_label.grid(row=7, column=0, sticky="e", padx=10, pady=10)
		
		pn_entry = tk.Entry(t, width=20, textvariable=varList[7])
		pn_entry.grid(row=7, column=1, sticky="w", padx=10, pady=10)
		
		preservation_label = tk.Label(t, text="Preservation Method:")
		preservation_label.grid(row=8, column=0, sticky="e", padx=10, pady=10)
		
		preservation_entry = tk.Entry(t, width=5, textvariable=varList[8])
		preservation_entry.grid(row=8, column=1, sticky="w", padx=10, pady=10)
		
		date_label = tk.Label(t, text="Due Date:")
		date_label.grid(row=9, column=0, sticky="e", padx=10, pady=10)
		
		date_entry = tk.Entry(t, width=12, textvariable=varList[9])
		date_entry.grid(row=9, column=1, sticky="w", padx=10, pady=10)
		
		submit_button = tk.Button(t, text="Submit Data", command=lambda: self.process_addition(varList,t))
		submit_button.grid(row=10, column=1, sticky="w", padx=10, pady=10)
		
	def contract_window(self):
		main_window = tk.Toplevel()
		main_window.geometry('300x200')
		main_window.title('Contract Management')
		
		t_path = Path('config_dict.json')
		if t_path.is_file():
			self.dict = json.load(open('config_dict.json'))
			self.main_wb = openpyxl.load_workbook(self.dict['main'])
			self.wip_wb = openpyxl.load_workbook(self.dict['wip'])
		
		add_button = tk.Button(main_window, text="Add New Contract", command= self.add_contract)
		add_button.grid(row=0, column=0, padx=10, pady=10)
		
		
		