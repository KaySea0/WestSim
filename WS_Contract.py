import tkinter as tk
from tkinter import ttk
import json
import openpyxl
import os
from pathlib import Path
from tkinter import filedialog
from myimages import *

# if no preservation method is given / is special mil spec, put "n/a" for method in display
# maybe add way to select multiple contracts that apply to certain PO

class WS_Contract(object):

	def __init__(self):
		self.main_wb = None
		self.wip_wb = None
		self.dict = None
		
		self.PO_dict = None
		self.company_list = None
		self.current_company = tk.StringVar()
		
		self.wip_dict = None
		self.wip_list = None
		self.current_contract = tk.StringVar()
		
		self.PO_Vars = None
		self.check_var = tk.IntVar()
		self.po_num = None
		
	def create_dicts(self):
		cage_wb = openpyxl.load_workbook(self.dict['cage'])
		cage_ws = cage_wb.active

		self.PO_dict = {}
		self.company_list = []
		for row in cage_ws.iter_rows(min_row=2, values_only=True):
			t_entry = {}
			t_entry['line1'] = row[5] if row[5] != None else ""
			t_entry['line2'] = row[6] if row[6] != None else ""
			t_entry['line3'] = row[7] if row[7] != None else ""
			t_entry['line4'] = row[8] if row[8] != None else ""
			t_entry['line5'] = row[9] if row[9] != None else ""
			
			self.company_list.append(row[2])
			self.PO_dict[row[2]] = t_entry
			
		wip_ws = self.wip_wb.active
		self.wip_dict = {}
		self.wip_list = []
		for row in wip_ws.iter_rows(min_row=2, values_only=True):
			t_entry = {}
			t_entry['pn'] = row[7]
			t_entry['nsn'] = row[4]
			t_entry['description'] = row[5]
			t_entry['qty'] = row[2]
			
			self.wip_list.append(row[1])
			self.wip_dict[row[1]] = t_entry
			
	def save_PO(self):
		f = filedialog.asksaveasfile(mode='w', initialfile="PurchaseOrder - " + self.PO_Vars[5].get(), filetypes=(("Excel file", "*.xlsx"),))
		
		po_template = openpyxl.load_workbook('PO_Template.xlsx')
		po_ws = po_template.active
		
		try: 
			ws_image = openpyxl.drawing.image.Image('logo.png')
			ws_image.anchor(po_ws.cell('B2'))
			po_ws.add_image(ws_image)
		except:
			pass
		
		po_ws['B10'] = self.PO_Vars[0].get()
		po_ws['B11'] = self.PO_Vars[1].get()
		po_ws['B12'] = self.PO_Vars[2].get()
		po_ws['B13'] = self.PO_Vars[3].get()
		po_ws['B14'] = self.PO_Vars[4].get()
		po_ws['G10'] = "PO: {}".format(self.PO_Vars[5].get())
		po_ws['G12'] = "Quote: {}".format(self.PO_Vars[6].get())
		po_ws['G13'] = "Delivery: {}".format(self.PO_Vars[7].get())
		po_ws['G14'] = "Terms: {}".format(self.PO_Vars[8].get())
		po_ws['C18'] = "P/N: {}".format(self.PO_Vars[9].get())
		po_ws['C19'] = "NSN: {}".format(self.PO_Vars[10].get())
		po_ws['C20'] = self.PO_Vars[11].get()
		po_ws['E18'] = self.PO_Vars[12].get()
		po_ws['G18'] = self.PO_Vars[13].get()
		
		if self.check_var.get() == 1:
			po_ws['B46'] = "UPS Ground account no: 2Y642X"
			
		po_template.save(f.name+".xlsx")
		f.close()
		os.remove(f.name)
		
	def process_addition(self, var_list, window):
		main_ws = self.main_wb['DLAORDERS']
		next_row = main_ws.max_row+1
		
		main_ws['G'+str(next_row)] = var_list[0].get()
		main_ws['B'+str(next_row)] = var_list[1].get()
		main_ws['E'+str(next_row)] = var_list[2].get()
		main_ws['K'+str(next_row)] = var_list[3].get()
		main_ws['D'+str(next_row)] = var_list[4].get()
		main_ws['F'+str(next_row)] = var_list[6].get()
		main_ws['H'+str(next_row)] = var_list[10].get()
		
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
		wip_ws['K'+str(next_row)] = var_list[10].get()
		
		self.wip_wb.save(self.dict['wip'])
		
		window.destroy()
		
	def create_PO(self):
		t = tk.Toplevel()
		t.geometry('640x450')
		t.title("PO Creation")
		
		self.current_company.set("")
		self.current_contract.set("")
		
		self.PO_Vars = []
		for i in range(0,14):
			temp = tk.StringVar()
			self.PO_Vars.append(temp)
			
		def company_function(eventObject):
			company_info = self.PO_dict[self.current_company.get()]
			self.PO_Vars[0].set(company_info['line1'])
			self.PO_Vars[1].set(company_info['line2'])
			self.PO_Vars[2].set(company_info['line3'])
			self.PO_Vars[3].set(company_info['line4'])
			self.PO_Vars[4].set(company_info['line5'])
			
		def contract_function(eventObject):
			contract_info = self.wip_dict[self.current_contract.get()]
			self.PO_Vars[9].set(contract_info['pn'])
			self.PO_Vars[10].set(contract_info['nsn'])
			self.PO_Vars[11].set(contract_info['description'])
			self.PO_Vars[12].set(contract_info['qty'])
		
		company_menu = ttk.Combobox(t, textvariable=self.current_company, values=self.company_list)
		company_menu.bind('<<ComboboxSelected>>', company_function)
		company_menu.grid(row=0, column=0, columnspan=2, sticky="new", padx=10, pady=10)
		
		company_label = tk.Label(t, text="Vendor Name:")
		company_label.grid(row=1, column=0, sticky="w", padx=10, pady=10)
		
		company_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[0])
		company_entry.grid(row=1, column=1, sticky="w", padx=10, pady=10)
		
		addr1_label = tk.Label(t, text="Address Line 1:")
		addr1_label.grid(row=2, column=0, sticky="w", padx=10, pady=10)
		
		addr1_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[1])
		addr1_entry.grid(row=2, column=1, sticky="w", padx=10, pady=10)
		
		addr2_label = tk.Label(t, text="Address Line 2:")
		addr2_label.grid(row=3, column=0, sticky="w", padx=10, pady=10)
		
		addr2_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[2])
		addr2_entry.grid(row=3, column=1, sticky="w", padx=10, pady=10)
		
		phone_label = tk.Label(t, text="Phone:")
		phone_label.grid(row=4, column=0, sticky="w", padx=10, pady=10)
		
		phone_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[3])
		phone_entry.grid(row=4, column=1, sticky="w", padx=10, pady=10)
		
		attention_label = tk.Label(t, text="Attention:")
		attention_label.grid(row=5, column=0, sticky="w", padx=10, pady=10)
		
		attention_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[4])
		attention_entry.grid(row=5, column=1, sticky="w", padx=10, pady=10)
		
		UPS_check = tk.Checkbutton(t, text="Include UPS Ground account #", variable=self.check_var)
		UPS_check.grid(row=0, column=2, columnspan=2, padx=10, pady=10)
		
		poNum_label = tk.Label(t, text="PO #:")
		poNum_label.grid(row=1, column=2, sticky="w", padx=10, pady=10)
		
		poNum_entry = tk.Entry(t, width=10, textvariable=self.PO_Vars[5])
		poNum_entry.grid(row=1, column=3, sticky="w", padx=10, pady=10)
		
		quote_label = tk.Label(t, text="Quote:")
		quote_label.grid(row=2, column=2, sticky="w", padx=10, pady=10)
		
		quote_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[6])
		quote_entry.grid(row=2, column=3, sticky="w", padx=10, pady=10)
		
		delivery_label = tk.Label(t, text="Delivery:")
		delivery_label.grid(row=3, column=2, sticky="w", padx=10, pady=10)
		
		delivery_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[7])
		delivery_entry.grid(row=3, column=3, sticky="w", padx=10, pady=10)
		
		terms_label = tk.Label(t, text="Terms:")
		terms_label.grid(row=4, column=2, sticky="w", padx=10, pady=10)
		
		terms_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[8])
		terms_entry.grid(row=4, column=3, sticky="w", padx=10, pady=10)
		
		unit_label = tk.Label(t, text="Unit Price:")
		unit_label.grid(row=5, column=2, sticky="w", padx=10, pady=10)
		
		unit_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[13])
		unit_entry.grid(row=5, column=3, sticky="w", padx=10, pady=10)
		
		contract_menu = ttk.Combobox(t, textvariable=self.current_contract, values=self.wip_list)
		contract_menu.bind('<<ComboboxSelected>>', contract_function)
		contract_menu.grid(row=6, column=0, columnspan=2, sticky="new", padx=10, pady=10)
		
		pn_label = tk.Label(t, text="P/N:")
		pn_label.grid(row=7, column=0, sticky="w", padx=10, pady=10)
		
		pn_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[9])
		pn_entry.grid(row=7, column=1, sticky="w", padx=10, pady=10)
		
		nsn_label = tk.Label(t, text="NSN:")
		nsn_label.grid(row=8, column=0, sticky="w", padx=10, pady=10)
		
		nsn_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[10])
		nsn_entry.grid(row=8, column=1, sticky="w", padx=10, pady=10)
		
		descr_label = tk.Label(t, text="Part Description:")
		descr_label.grid(row=7, column=2, sticky="w", padx=10, pady=10)
		
		descr_entry = tk.Entry(t, width=30, textvariable=self.PO_Vars[11])
		descr_entry.grid(row=7, column=3, sticky="w", padx=10, pady=10)
		
		qty_label = tk.Label(t, text="QTY:")
		qty_label.grid(row=8, column=2, sticky="w", padx=10, pady=10)
		
		qty_entry=tk.Entry(t, width=30, textvariable=self.PO_Vars[12])
		qty_entry.grid(row=8, column=3, sticky="w", padx=10, pady=10)
		
		submit_btn = tk.Button(t, text="Create PO", command=self.save_PO)
		submit_btn.grid(row=9, column=3, sticky="e", padx=10, pady=10)
		
	def add_contract(self):
		t = tk.Toplevel()
		t.geometry('350x500')
		t.title("Add New Contract")
		
		varList = []
		for i in range(0,11):
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
		
		unit_label = tk.Label(t, text="Unit Price:")
		unit_label.grid(row=8, column=0, sticky="e", padx=10, pady=10)
		
		unit_entry = tk.Entry(t, width=10, textvariable=varList[8])
		unit_entry.grid(row=8, column=1, sticky="w", padx=10, pady=10)
		
		preservation_label = tk.Label(t, text="Preservation Method:")
		preservation_label.grid(row=9, column=0, sticky="e", padx=10, pady=10)
		
		preservation_entry = tk.Entry(t, width=5, textvariable=varList[9])
		preservation_entry.grid(row=9, column=1, sticky="w", padx=10, pady=10)
		
		date_label = tk.Label(t, text="Due Date:")
		date_label.grid(row=10, column=0, sticky="e", padx=10, pady=10)
		
		date_entry = tk.Entry(t, width=12, textvariable=varList[10])
		date_entry.grid(row=10, column=1, sticky="w", padx=10, pady=10)
		
		submit_button = tk.Button(t, text="Submit Data", command=lambda: self.process_addition(varList,t))
		submit_button.grid(row=11, column=1, sticky="w", padx=10, pady=10)
		
	def contract_window(self):
		main_window = tk.Toplevel()
		main_window.geometry('300x200')
		main_window.title('Contract Management')
		
		t_path = Path('config_dict.json')
		if t_path.is_file():
			self.dict = json.load(open('config_dict.json'))
			self.main_wb = openpyxl.load_workbook(self.dict['main'])
			self.wip_wb = openpyxl.load_workbook(self.dict['wip'])
			self.create_dicts()
		
		add_button = tk.Button(main_window, text="Add New Contract", command= self.add_contract)
		add_button.grid(row=0, column=0, padx=10, pady=10)
		
		create_PO_button = tk.Button(main_window, text="Create PO", command= self.create_PO)
		create_PO_button.grid(row=1, column=0, padx=10, pady=10)
		
		
		