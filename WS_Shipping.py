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
		
	def save_changes(self):
		pass
	
	# start at 251
	def create_lists(self):
		main_ws = self.main_wb['DLAORDERS']
		
		
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