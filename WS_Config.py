import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import json

class WS_Config(object):

	def __init__(self):
		
		self.config_dict = json.load(open('config_dict.json'))
			
		self.bid_var = tk.StringVar()
		self.bid_var.set(self.config_dict["bid"])
		
	def bid_browse(self):
		self.config_dict['bid'] = filedialog.askdirectory()
		self.bid_var.set(self.config_dict["bid"])
		
	def save_config(self,t):
		json_temp = json.dumps(self.config_dict)
		f = open("config_dict.json","w")
		f.write(json_temp)
		f.close()
		
		messagebox.showinfo("Config Confirmation", "Your changes have been saved!")
		t.destroy()
		
	def config_window(self):
		
		t = tk.Toplevel()
		t.title("Configuration Window")
		t.geometry('600x200')
		
		bid_label = tk.Label(t, text="Bid Sheet Folder")
		bid_label.grid(row=0, column=0, padx=10, pady=10)
		
		bid_entry = tk.Entry(t, state="disabled", textvar= self.bid_var, width=60)
		bid_entry.grid(row=0, column=1, padx=10, pady=10)
		
		bid_browse = tk.Button(t, text="Browse", command= self.bid_browse)
		bid_browse.grid(row=0, column=2, padx=10, pady=10)
		
		save_button = tk.Button(t, text="Save Changes", command=lambda: self.save_config(t))
		save_button.grid(row=1,column=0,pady=10)