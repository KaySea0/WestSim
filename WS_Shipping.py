import tkinter as tk
from pathlib import Path
from tkinter import messagebox

# run through main_wb and grab PO number and contract number for each line
# create another "search" window by PO number and contract number
# when selecting contract, have two buttons for different popups
# # labels for initial searching in VSM (first 6 / last 4 of contract number), label info (PN, Preservation Method), and company name for label saving
# # input for receiving report (toggle for total/partial, quantity, shipping/invoice number, total for verification, RFID number) with toggle for RFID

class WS_Shipping(object):
	
	def __init__(self):
	
		self.contract_list = None
		self.wip_list = None
		
		self.dict = None
		self.main_wb = None
		self.wip_wb = None
		
		self.contract_var = tk.StringVar()
		self.po_var = tk.StringVar()
		
		self.search_frame = None
		self.canvas = None
		self.scroll = None
		
	def update_view(self, *args):
		pass
		
	def save_changes(self):
		pass
	
	def create_lists(self):
		main_ws = self.main_wb['DLAORDERS']
		list_start = 278 # row number that search should start at for main_ws (dependent on what orders have come in)
		list_end = main_ws.max_row+1
		
		# Contract Number - PO Number - Vendor Name
		for i in range(list_start, list_end):
			if not main_ws['I'+str(i)] is None:
				self.contract_list.append([main_ws['B'+str(i)], main_ws['I'+str(i)], main_ws['F'+str(i)]])
				
		wip_ws = self.wip_wb.active
		list_end = wip_ws.max_row+1
		
		# Contract Number - Vendor Name - P/N - Preservation Method
		for i in range(1, list_end):
			self.wip_list.append([wip_ws['B'+str(i)], wip_ws['G'+str(i)], wip_ws['H'+str(i)], wip_ws['I'+str(i)]])
		
	def shipping_window(self):
		
		t = tk.TopLevel()
		t.geometry('250x150')
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
			
		self.contract_var.trace("w", self.update_view)
		self.po_var.trace("w", self.update_view)
			
		self.canvas = tk.Canvas(t, borderwidth=0)
		self.scroll = tk.Scrollbar(t, orient="vertical", command=self.canvas.yview)
		
		input_frame = tk.Frame(t)
		
		contract_label = tk.Label(input_frame, text="Contract Number:")
		contract_label.grid(row=0, column=0, padx=5, pady=5)
		
		contract_entry = tk.Entry(input_frame, width=30, textvariable=self.contract_var)
		contract_entry.grid(row=0, column=1, padx=5, pady=5)
		
		po_label = tk.Label(input_frame, text="PO Number:")
		po_label.grid(row=0, column=2, padx=5, pady=5)
		
		po_entry = tk.Entry(input_frame, width=10, textvariable=self.po_var)
		po_entry.grid(row=0, column=3, padx=5, pady=5)
		
		input_frame.pack(anchor="n", fill=tk.X)
		self.update_view()
			