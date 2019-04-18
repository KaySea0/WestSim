# https://stackoverflow.com/questions/27745500/how-to-save-a-list-to-a-file-and-read-it-as-a-list-type/28305183
import tkinter as tk
import pickle
from tkinter import messagebox

# record part number and NSN for each part
# PN ---- NSN ---- QTY ---- Edit
# Edit -> message box asking for qty change -> change value in list
# Save function that is used when editing and adding

class WS_Inventory(object):
	
	def __init__(self):
		
		try:
			with open("inventory.txt","rb") as fp:
				self.inventory_list = pickle.load(fp)
		except:
			self.inventory_list = None
			
		self.pn_var = tk.StringVar()
		self.nsn_var = tk.StringVar()
		self.qty_var = tk.IntVar()
			
	def search_window(self):
		sw = tk.Toplevel()
		sw.geometry('400x300')
		sw.title("Search Inventory")
		
		pn_label = tk.Label(sw, text="PN:")
		pn_label.grid(row=0, column=0, padx=5, pady=5)
		
		pn_entry = tk.Entry(sw, width=25, textvariable=self.pn_var)
		pn_entry.grid(row=0, column=1, padx=5, pady=5)
		
		nsn_label = tk.Label(sw, text="NSN:")
		nsn_label.grid(row=0, column=2, padx=5, pady=5)
		
		nsn_entry = tk.Entry(sw, width=15, textvariable=self.nsn_var)
		nsn_entry.grid(row=0, column=3, padx=5, pady=5)
		
		search_button = tk.Button(sw, text="Search")
		search_button.grid(row=0, column=4, padx=5, pady=5)
		
	def add_stock(self):
		
		self.inventory_list.append([len(self.inventory_list)+1, self.pn_var.get(), self.nsn_var.get(), self.qty_var.get()])
		
		self.pn_var.set('')
		self.nsn_var.set('')
		self.qty_var.set(0)
		
		messagebox.showinfo("Stock Added", "Stock information has been saved!")
		
		
	def add_stock_window(self):
		win = tk.Toplevel()
		win.geometry('250x130')
		win.title("Add Stock")
		
		pn_label = tk.Label(win, text="PN:")
		pn_label.grid(row=0, column=0, padx=5, pady=5)
		
		pn_entry = tk.Entry(win, width=25, textvariable=self.pn_var)
		pn_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
		
		nsn_label = tk.Label(win, text="NSN:")
		nsn_label.grid(row=1, column=0, padx=5, pady=5)
		
		nsn_entry = tk.Entry(win, width=15, textvariable=self.nsn_var)
		nsn_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
		
		qty_label = tk.Label(win, text="Quantity:")
		qty_label.grid(row=2, column=0, padx=5, pady=5)
		
		qty_entry = tk.Entry(win, width=5, textvariable=self.qty_var)
		qty_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
		
		submit_button = tk.Button(win, text="Submit")
		submit_button.grid(row=3, column=1, padx=5, pady=5, sticky="w")
		
	def inventory_window(self):
		t = tk.Toplevel()
		t.geometry('250x200')
		t.title("Inventory Management")
		
		search_button = tk.Button(t, text="Search Inventory", command=self.search_window)
		search_button.grid(row=0, column=0, padx=10, pady=10)
		
		add_button = tk.Button(t, text="Add Stock", command=self.add_stock_window)
		add_button.grid(row=1, column=0, padx=10, pady=10)
		