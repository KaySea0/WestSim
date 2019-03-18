import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import json
from pathlib import Path
from dict_init import cage_dict_init

# max length of part name = 15

class WS_Config(object):

	def __init__(self):
		
		self.bid_var = tk.StringVar()
		self.cage_var = tk.StringVar()
		self.main_var = tk.StringVar()
		self.contract_var = tk.StringVar()
		t_path = Path('config_dict.json')
		
		if t_path.is_file():
			self.config_dict = json.load(open('config_dict.json'))
			self.bid_var.set(self.config_dict["bid"])
			self.cage_var.set(self.config_dict["cage"])
			self.main_var.set(self.config_dict["main"])
			self.contract_var.set(self.config_dict["contract"])
		else:
			self.config_dict = {}
			
	def bid_browse(self):
		self.config_dict['bid'] = filedialog.askdirectory()
		self.bid_var.set(self.config_dict["bid"])
		
	def cage_browse(self):
		self.config_dict['cage'] = filedialog.askopenfilename()
		self.cage_var.set(self.config_dict['cage'])
		
	def main_browse(self):
		self.config_dict['main'] = filedialog.askopenfilename()
		self.main_var.set(self.config_dict['main'])
		
	def contract_browse(self):
		self.config_dict['contract'] = filedialog.askopenfilename()
		self.contract_var.set(self.config_dict['contract'])
		
	def save_config(self,t):
		json_temp = json.dumps(self.config_dict)
		f = open("config_dict.json","w")
		f.write(json_temp)
		f.close()
		
		messagebox.showinfo("Config Confirmation", "Configuration settings have been saved!")
		t.destroy()
		
	def config_window(self):
		
		t = tk.Toplevel()
		t.title("Configuration Window")
		t.geometry('600x250')
		
		bid_label = tk.Label(t, text="Bid Sheet Folder")
		bid_label.grid(row=0, column=0, padx=10, pady=10)
		
		bid_entry = tk.Entry(t, state="disabled", textvar= self.bid_var, width=60)
		bid_entry.grid(row=0, column=1, padx=10, pady=10)
		
		bid_browse = tk.Button(t, text="Browse", command= self.bid_browse)
		bid_browse.grid(row=0, column=2, padx=10, pady=10)
		
		cage_label = tk.Label(t, text="Cagecode Workbook")
		cage_label.grid(row=1, column=0, padx=10, pady=10)
		
		cage_entry = tk.Entry(t, state="disabled", textvar = self.cage_var, width=60)
		cage_entry.grid(row=1, column=1, padx=10, pady=10)
		
		cage_browse = tk.Button(t, text="Browse", command= self.cage_browse)
		cage_browse.grid(row=1, column=2, padx=10, pady=10)
		
		main_label = tk.Label(t, text="Main Workbook")
		main_label.grid(row=2, column=0, padx=10, pady=10)
		
		main_entry = tk.Entry(t, state="disabled", textvar = self.main_var, width=60)
		main_entry.grid(row=2, column=1, padx=10, pady=10)
		
		main_browse = tk.Button(t, text="Browse", command= self.main_browse)
		main_browse.grid(row=2, column=2, padx=10, pady=10)
		
		contract_label = tk.Label(t, text="Contract Workbook")
		contract_label.grid(row=3, column=0, padx=10, pady=10)
		
		contract_entry = tk.Entry(t, state="disabled", textvar = self.contract_var, width=60)
		contract_entry.grid(row=3, column=1, padx=10, pady=10)
		
		contract_browse = tk.Button(t, text="Browse", command= self.contract_browse)
		contract_browse.grid(row=3, column=2, padx=10, pady=10)
		
		cage_dict_create = tk.Button(t, text="Process Cagecodes", command=lambda: cage_dict_init(self.config_dict['cage']))
		cage_dict_create.grid(row=4, column=0, padx=10, pady=10)
		
		save_button = tk.Button(t, text="Save Changes", command=lambda: self.save_config(t))
		save_button.grid(row=4, column=1, pady=10)