import tkinter as tk
import json
import openpyxl
from pathlib import Path

# if no preservation method is given / is special mil spec, put "n/a" for method in display

class WS_Contract(object):

	def __init__(self):
		self.main_wb = None
		self.wip_wb = None
		
	def add_contract(self):
		t = tk.Toplevel()
		t.geometry('300x500')
		t.title("Add New Contract")
		
		date_label = tk.Label(t, text="Data Awarded:")
		date_label.grid(row=0, column=0, sticky="e", padx=10, pady=10)
		
		date_entry = tk.Entry(t, width=12)
		date_entry.grid(row=0, column=1, sticky="w", padx=10, pady=10)
		
		cnumber_label = tk.Label(t, text="Contract #:")
		cnumber_label.grid(row=1, column=0, sticky="e", padx=10, pady=10)
		
		cnumber_entry = tk.Entry(t, width=18)
		cnumber_entry.grid(row=1, column=1, sticky="w", padx=10, pady=10)
		
		qty_label = tk.Label(t, text="Quantity:")
		qty_label.grid(row=2, column=0, sticky="e", padx=10, pady=10)
		
		qty_entry = tk.Entry(t, width=5)
		qty_entry.grid(row=2, column=1, sticky="w", padx=10, pady=10)
		
		ctotal_label = tk.Label(t, text="Contract Total:")
		ctotal_label.grid(row=3, column=0, sticky="e", padx=10, pady=10)
		
		ctotal_entry = tk.Entry(t, width=11)
		ctotal_entry.grid(row=3, column=1, sticky="w", padx=10, pady=10)
		
		nsn_label = tk.Label(t, text="NSN:")
		nsn_label.grid(row=4, column=0, sticky="e", padx=10, pady=10)
		
		nsn_entry = tk.Entry(t, width=15)
		nsn_entry.grid(row=4, column=1, sticky="w", padx=10, pady=10)
		
		vendor_label = tk.Label(t, text="Vendor Name:")
		vendor_label.grid(row=5, column=0, sticky="e", padx=10, pady=10)
		
		vendor_entry = tk.Entry(t, width=25)
		vendor_entry.grid(row=5, column=1, sticky="w", padx=10, pady=10)
		
		pn_label = tk.Label(t, text="Part #:")
		pn_label.grid(row=6, column=0, sticky="e", padx=10, pady=10)
		
		pn_entry = tk.Entry(t, width=20)
		pn_entry.grid(row=6, column=1, sticky="w", padx=10, pady=10)
		
		preservation_label = tk.Label(t, text="Preservation Method:")
		preservation_label.grid(row=7, column=0, sticky="e", padx=10, pady=10)
		
		preservation_entry = tk.Entry(t, width=5)
		preservation_entry.grid(row=7, column=1, sticky="w", padx=10, pady=10)
		
		date_label = tk.Label(t, text="Due Date:")
		date_label.grid(row=8, column=0, sticky="e", padx=10, pady=10)
		
		date_entry = tk.Entry(t, width=12)
		date_entry.grid(row=8, column=1, sticky="w", padx=10, pady=10)
		
	def contract_window(self):
		main_window = tk.Toplevel()
		main_window.geometry('300x200')
		main_window.title('Contract Management')
		
		t_path = Path('config_dict.json')
		if t_path.is_file():
			t_dict = json.load(open('config_dict.json'))
			self.main_wb = t_dict['main']
			self.wip_wb = t_dict['contract']
		
		add_button = tk.Button(main_window, text="Add New Contract", command= self.add_contract)
		add_button.grid(row=0, column=0, padx=10, pady=10)
		
		
		