import tkinter as tk
from pathlib import Path
from tkinter import messagebox
import json
import openpyxl
import os
import datetime

# run through main_wb and grab PO number and contract number for each line
# create another "search" window by PO number and contract number
# when selecting contract, have two buttons for different popups
# # labels for initial searching in VSM (first 6 / last 4 of contract number), label info (PN, Preservation Method), and company name for label saving
# # input for receiving report (toggle for total/partial, quantity, shipping/invoice number, total for verification, RFID number) with toggle for RFID

class WS_Shipping(object):
	
	def __init__(self):
	
		self.contract_list = []
		self.wip_list = []
		
		self.dict = None
		self.main_wb = None
		self.wip_wb = None
		
		self.contract_var = tk.StringVar()
		self.po_var = tk.StringVar()
		
		self.search_frame = None
		self.canvas = None
		self.scroll = None
		
	def update_view(self, *args):
		
		if self.search_frame is not None:
			self.search_frame.destroy()
			
		self.search_frame = tk.Frame(self.canvas)
		
		contract_search_label = tk.Label(self.search_frame, text="Contract #")
		contract_search_label.grid(row=0, column=0, padx=5, pady=5)
		
		po_search_label = tk.Label(self.search_frame, text="PO #")
		po_search_label.grid(row=0, column=1, padx=5, pady=5)
		
		vendor_search_label = tk.Label(self.search_frame, text="Vendor")
		vendor_search_label.grid(row=0, column=2, padx=5, pady=5)
		
		if not self.contract_var.get() and not self.po_var.get():
			display_list = self.contract_list
		else: 
			
			display_list = []
			
			for contract in self.contract_list:
				if self.contract_var.get():
					if self.po_var.get():
						if self.contract_var.get().lower() in contract[0].lower() and self.po_var.get().lower() in contract[1].lower():
							display_list.append(contract)
							
					else:
						if self.contract_var.get().lower() in contract[0].lower():
							display_list.append(contract)
							
				else:
					if self.po_var.get().lower() in contract[1].lower():
						display_list.append(contract)
						
		for i in range(1, len(display_list)+1):
			
			contract = display_list[i-1]
			
			contract_single_label = tk.Label(self.search_frame, text=contract[0])
			contract_single_label.grid(row=i, column=0, padx=5, pady=5)
			
			po_single_label = tk.Label(self.search_frame, text=contract[1])
			po_single_label.grid(row=i, column=1, padx=5, pady=5)
			
			vendor_single_label = tk.Label(self.search_frame, text=contract[2])
			vendor_single_label.grid(row=i, column=2, padx=5, pady=5)
			
			vsm_button = tk.Button(self.search_frame, text="VSM", command=lambda contract=contract: self.vsm_window(contract))
			vsm_button.grid(row=i, column=3, padx=5, pady=5)
			
			wawf_button = tk.Button(self.search_frame, text="WAWF", command=lambda contract=contract: self.wawf_window(contract))
			wawf_button.grid(row=i, column=4, padx=5, pady=5)
			
		self.canvas.create_window((4,4),window=self.search_frame,anchor="nw")
		
		self.canvas.update_idletasks()
		self.canvas.configure(scrollregion=self.canvas.bbox('all'), yscrollcommand=self.scroll.set)
		
		self.canvas.pack(side="left",fill="both",expand=True)
		self.scroll.pack(side="right",fill="y")
		
	def vsm_window(self, contract):
		vsm = tk.Toplevel()
		vsm.title('VSM - ' + contract[0])
		
		company_name = tk.Label(vsm, text=contract[2], font='Helvetica 15 bold')
		company_name.grid(row=0, column=0, sticky="e", padx=5, pady=5)
		
		if isinstance(contract[3], str):
			date_text = contract[3]
		else:
			date_text = contract[3].strftime('%m/%d/%Y')
		
		contract_date = tk.Label(vsm, text=date_text, font='Helvetica 15 bold')
		contract_date.grid(row=0, column=1, sticky="w", padx=5, pady=5)
		
		po_label = tk.Label(vsm, text="PO Search:")
		po_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
		
		po_search = tk.Text(vsm, height=1, width=8, borderwidth=0)
		po_search.insert(1.0, contract[0][:6])
		po_search.configure(state="disabled", inactiveselectbackground=po_search.cget("selectbackground"))
		po_search.grid(row=1, column=1, sticky="w", padx=5, pady=5)
		
		po_ref = tk.Label(vsm, text="-" + contract[0][contract[0].rfind('-')+1:])
		po_ref.grid(row=1, column=2, padx=5, pady=5)
		
		pn_text = ""
		preservation_text = ""
		
		for ref in self.wip_list:
			if ref[0] == contract[0]:
				pn_text = ref[2]
				preservation_text = ref[3]
				
		if not pn_text: 
			pn_text = "Look up in contract PDF"
			preservation_text = "Look up in contract PDF"
			
		pn_label = tk.Label(vsm, text="P/N:")
		pn_label.grid(row=2, column=0, sticky="e", padx=5, pady=5)
		
		pn_info = tk.Text(vsm, height=1, width=20, borderwidth=0)
		pn_info.insert(1.0, pn_text)
		pn_info.configure(state="disabled", inactiveselectbackground=po_search.cget("selectbackground"))
		pn_info.grid(row=2, column=1, sticky="w", padx=5, pady=5)
		
		preservation_label = tk.Label(vsm, text="Preservation Method:")
		preservation_label.grid(row=3, column=0, sticky="e", padx=5, pady=5)
		
		preservation_info = tk.Text(vsm, height=1, width=20, borderwidth=0)
		preservation_info.insert(1.0, preservation_text)
		preservation_info.configure(state="disabled", inactiveselectbackground=po_search.cget("selectbackground"))
		preservation_info.grid(row=3, column=1, sticky="w", padx=5, pady=5)
		
	def wawf_window(self, contract):
		pass
	
	def save_changes(self):
		pass
	
	def create_lists(self):
	
		if not self.contract_list:
			main_ws = self.main_wb['DLAORDERS']
			list_start = 278 # row number that search should start at for main_ws (dependent on what orders have come in)
			list_end = main_ws.max_row+1
			
			# Contract Number - PO Number - Vendor Name - Date Awarded
			for i in range(list_start, list_end):
				if not main_ws['I'+str(i)].value is None:
					self.contract_list.append([main_ws['B'+str(i)].value, main_ws['I'+str(i)].value, main_ws['F'+str(i)].value, main_ws['G'+str(i)].value])
					
			wip_ws = self.wip_wb.active
			list_end = wip_ws.max_row+1
			
			# Contract Number - Vendor Name - P/N - Preservation Method
			for i in range(1, list_end):
				self.wip_list.append([wip_ws['B'+str(i)].value, wip_ws['G'+str(i)].value, wip_ws['H'+str(i)].value, wip_ws['I'+str(i)].value])
				
	def shipping_window(self):
		
		t = tk.Toplevel()
		t.geometry('525x250')
		t.title('Shipping Management')
		
		def _delete_window():
			try:
				self.save_changes()
				t.destroy()
			except:
				pass
				
		t.protocol("WM_DELETE_WINDOW", _delete_window)
		
		t_path = Path('config_dict.json')
		if t_path.is_file():
			self.dict = json.load(open('config_dict.json'))
			self.main_wb = openpyxl.load_workbook(self.dict['main'])
			self.wip_wb = openpyxl.load_workbook(self.dict['wip'])
			self.create_lists()
			
		self.canvas = tk.Canvas(t, borderwidth=0)
		self.scroll = tk.Scrollbar(t, orient="vertical", command=self.canvas.yview)
		
		input_frame = tk.Frame(t)
		
		contract_label = tk.Label(input_frame, text="Contract Number:")
		contract_label.grid(row=0, column=0, padx=5, pady=5)
		
		contract_entry = tk.Entry(input_frame, width=20, textvariable=self.contract_var)
		contract_entry.grid(row=0, column=1, padx=5, pady=5)
		
		po_label = tk.Label(input_frame, text="PO Number:")
		po_label.grid(row=0, column=2, padx=5, pady=5)
		
		po_entry = tk.Entry(input_frame, width=10, textvariable=self.po_var)
		po_entry.grid(row=0, column=3, padx=5, pady=5)
		
		search_button = tk.Button(input_frame, text="Search", command=self.update_view)
		search_button.grid(row=0, column=4, padx=5, pady=5)
		
		input_frame.pack(anchor="n", fill=tk.X)
		