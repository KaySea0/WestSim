import tkinter as tk
import pickle
from tkinter import messagebox
from pathlib import Path
import json

# change search function to work with ID once all stock has been marked

class WS_Inventory(object):
	
	def __init__(self):
		
		# load inventory list if it exists, create empty list otherwise
		t_path = Path('config_dict.json')
		if t_path.is_file():
			self.inv_path = json.load(open('config_dict.json'))['inv']
			try:
				with open(self.inv_path,"rb") as fp:
					self.inventory_list = pickle.load(fp)
			except:
				self.inventory_list = []
			
		self.pn_var = tk.StringVar()  # reference variable for part number
		self.nsn_var = tk.StringVar() # reference variable for NSN
		self.qty_var = tk.IntVar()    # reference variable for quantity
		self.id_var = tk.StringVar()  # reference variable for stock id
		
		self.pn_add_var = tk.StringVar()  # reference variable for part number - addition
		self.nsn_add_var = tk.StringVar() # reference variable for NSN - addition
		self.qty_add_var = tk.IntVar()    # reference variable for quantity - addition
		self.id_add_var = tk.StringVar()  # reference variable for stock id - addition
		
		self.search_frame = None # frame that contains list of stock that matches search
		self.canvas = None # canvas that contains search frame
		self.scroll = None # scrollbar that controls aforementioned list of stock
	
	# # #
	# Method: update_view
	# Input: n/a
	# Utility:
	#   Update displayed list of stock based on search terms (PN/NSN) provided by user in text boxes
	# # #
	def update_view(self, *args):
		
		# destroy current frame before rebuilding
		if self.search_frame is not None:
			self.search_frame.destroy()
			
		# create new frame
		self.search_frame = tk.Frame(self.canvas)
		
		# label for PN column
		pn_header = tk.Label(self.search_frame, text="PN")
		pn_header.grid(row=0, column=0, padx=10, pady=10)
		
		# label for NSN column
		nsn_header = tk.Label(self.search_frame, text="NSN")
		nsn_header.grid(row=0, column=1, padx=10, pady=10)
		
		# label for QTY column
		qty_header = tk.Label(self.search_frame, text="QTY")
		qty_header.grid(row=0, column=2, padx=10, pady=10)
		
		# label for ID column
		id_header = tk.Label(self.search_frame, text="ID")
		id_header.grid(row=0, column=3, padx=10, pady=10)
		
		# if neither text box has any data, display full inventory list
		if not self.pn_var.get() and not self.nsn_var.get():
			display_list = self.inventory_list
		else: # if either / both boxes are populated, perform search for matching stock
			
			# list to store matching stock that will be displayed
			display_list = []
			
			# running through every item currently in inventory...
			for part in self.inventory_list:
				
				# user looking for matching sequence in both PN and NSN
				if self.pn_var.get():
					if self.nsn_var.get():
						if self.pn_var.get().lower() in part[1].lower() and self.nsn_var.get().lower() in part[2].lower():
							display_list.append(part)
					# user looking for matching sequence in only PN
					else:
						if self.pn_var.get().lower() in part[1].lower():
							display_list.append(part)
				# user looking for matching sequence in only NSN
				else:
					if self.nsn_var.get().lower() in part[2].lower():
						display_list.append(part)
		
		# for matched stock, create item to display on search frame 
		for i in range(1, len(display_list)+1):
		
			# quick reference for current item
			part = display_list[i-1]
			
			# PN label for stock
			pn_search_label = tk.Label(self.search_frame, text=part[1])
			pn_search_label.grid(row=i, column=0, padx=10, pady=10)
			
			# NSN label for stock
			nsn_search_label = tk.Label(self.search_frame, text=part[2])
			nsn_search_label.grid(row=i, column=1, padx=10, pady=10)
			
			# QTY label for stock
			qty_search_label = tk.Label(self.search_frame, text=part[3])
			qty_search_label.grid(row=i, column=2, padx=10, pady=10)
			
			# ID label for stock
			try:
				id_search_label = tk.Label(self.search_frame, text=part[4])
			except IndexError:
				id_search_label = tk.Label(self.search_frame, text='')
			id_search_label.grid(row=i, column=3, padx=10, pady=10)
			
			# 'Edit' button for each item to change 
			edit_button = tk.Button(self.search_frame, text="Edit", command=lambda item=part: self.edit_window(item))
			edit_button.grid(row=i, column=4, padx=10, pady=10)
			
		self.canvas.create_window((4,4),window=self.search_frame,anchor="nw")
		
		self.canvas.update_idletasks()
		self.canvas.configure(scrollregion=self.canvas.bbox('all'), yscrollcommand=self.scroll.set)
		
		self.canvas.pack(side="left",fill="both",expand=True)
		self.scroll.pack(side="right",fill="y")
		
	def search_window(self):
		sw = tk.Toplevel()
		sw.geometry('450x300')
		sw.title("Search Inventory")
		
		self.canvas = tk.Canvas(sw, borderwidth=0)
		self.scroll = tk.Scrollbar(sw, orient="vertical", command=self.canvas.yview)
		
		input_frame = tk.Frame(sw)
		
		pn_label = tk.Label(input_frame, text="PN:")
		pn_label.grid(row=0, column=0, padx=5, pady=5)
		
		pn_entry = tk.Entry(input_frame, width=25, textvariable=self.pn_var)
		pn_entry.grid(row=0, column=1, padx=5, pady=5)
		
		nsn_label = tk.Label(input_frame, text="NSN:")
		nsn_label.grid(row=0, column=2, padx=5, pady=5)
		
		nsn_entry = tk.Entry(input_frame, width=20, textvariable=self.nsn_var)
		nsn_entry.grid(row=0, column=3, padx=5, pady=5)
		
		input_frame.pack(anchor="n", fill=tk.X)
		self.update_view()
		
	def add_stock(self, win):
		
		self.inventory_list.append([len(self.inventory_list)+1, self.pn_add_var.get(), self.nsn_add_var.get(), self.qty_add_var.get(), self.id_add_var.get()])
		
		self.pn_add_var.set('')
		self.nsn_add_var.set('')
		self.qty_add_var.set(0)
		self.id_add_var.set('')
		
		win.destroy()
		messagebox.showinfo("Stock Added", "Stock information has been saved!")
		
	def edit_stock(self, index, win):
		
		self.inventory_list[index][1] = self.pn_var.get()
		self.inventory_list[index][2] = self.nsn_var.get()
		self.inventory_list[index][3] = self.qty_var.get()
		try:
			self.inventory_list[index][4] = self.id_var.get()
		except IndexError:
			self.inventory_list[index].append(self.id_var.get())
		
		self.pn_var.set('')
		self.nsn_var.set('')
		self.qty_var.set(0)
		self.id_var.set('')
		
		win.destroy()
		messagebox.showinfo("Stock Edited", "Stock information has been saved!")
		
	def save_changes(self):
		
		with open(self.inv_path,"wb") as fp:
			pickle.dump(self.inventory_list, fp)
		
		messagebox.showinfo("Data Saved", "Changes to inventory have been saved!")
			
	def edit_window(self, item):
		e_window = tk.Toplevel()
		e_window.geometry('250x170')
		e_window.title("Edit Stock")
		
		self.pn_var.set(item[1])
		self.nsn_var.set(item[2])
		self.qty_var.set(item[3])
		try:
			self.id_var.set(item[4])
		except IndexError:
			self.id_var.set('')
		
		pn_label = tk.Label(e_window, text="PN:")
		pn_label.grid(row=0, column=0, padx=5, pady=5)
		
		pn_entry = tk.Entry(e_window, width=25, textvariable=self.pn_var)
		pn_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
		
		nsn_label = tk.Label(e_window, text="NSN:")
		nsn_label.grid(row=1, column=0, padx=5, pady=5)
		
		nsn_entry = tk.Entry(e_window, width=15, textvariable=self.nsn_var)
		nsn_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
		
		qty_label = tk.Label(e_window, text="Quantity:")
		qty_label.grid(row=2, column=0, padx=5, pady=5)
		
		qty_entry = tk.Entry(e_window, width=5, textvariable=self.qty_var)
		qty_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
		
		id_label = tk.Label(e_window, text="ID:")
		id_label.grid(row=3, column=0, padx=5, pady=5)
		
		id_entry = tk.Entry(e_window, width=5, textvariable=self.id_var)
		id_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
		
		submit_button = tk.Button(e_window, text="Submit", command=lambda: self.edit_stock(item[0]-1, e_window))
		submit_button.grid(row=4, column=1, padx=5, pady=5, sticky="w")
	
	def add_stock_window(self):
		win = tk.Toplevel()
		win.geometry('250x170')
		win.title("Add Stock")
		
		pn_label = tk.Label(win, text="PN:")
		pn_label.grid(row=0, column=0, padx=5, pady=5)
		
		pn_entry = tk.Entry(win, width=25, textvariable=self.pn_add_var)
		pn_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
		
		nsn_label = tk.Label(win, text="NSN:")
		nsn_label.grid(row=1, column=0, padx=5, pady=5)
		
		nsn_entry = tk.Entry(win, width=15, textvariable=self.nsn_add_var)
		nsn_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
		
		qty_label = tk.Label(win, text="Quantity:")
		qty_label.grid(row=2, column=0, padx=5, pady=5)
		
		qty_entry = tk.Entry(win, width=5, textvariable=self.qty_add_var)
		qty_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
		
		id_label = tk.Label(win, text="ID:")
		id_label.grid(row=3, column=0, padx=5, pady=5)
		
		id_entry = tk.Entry(win, width=5, textvariable=self.id_add_var)
		id_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
		
		submit_button = tk.Button(win, text="Submit", command=lambda: self.add_stock(win))
		submit_button.grid(row=4, column=1, padx=5, pady=5, sticky="w")
		
	def inventory_window(self):
		t = tk.Toplevel()
		t.geometry('275x75')
		t.title("Stock Management")
		
		def _delete_window():
			try:
				self.save_changes()
				t.destroy()
			except:
				pass
				
		t.protocol("WM_DELETE_WINDOW", _delete_window)
		
		self.pn_var.trace("w", self.update_view)
		self.nsn_var.trace("w", self.update_view)
		
		search_button = tk.Button(t, text="Search Stock", command=self.search_window)
		search_button.grid(row=0, column=0, padx=10, pady=5)
		
		add_button = tk.Button(t, text="Add Stock", command=self.add_stock_window)
		add_button.grid(row=0, column=2, padx=10, pady=5)
		
		save_button = tk.Button(t, text="Save Changes", command=self.save_changes)
		save_button.grid(row=1, column=1, padx=2, pady=5)
		