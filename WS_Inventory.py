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
			self.inventory_list = []
			
		self.pn_var = tk.StringVar()
		self.nsn_var = tk.StringVar()
		self.qty_var = tk.IntVar()
		
		self.search_frame = None
		self.canvas = None
		self.scroll = None
		
	def reset_scrollregion(self, event):
		self.canvas.configure(width=event.width, height=event.height, yscrollcommand=self.scroll.set)
		
	def update_view(self, *args):
		
		if self.search_frame is not None:
			self.search_frame.destroy()
			
		self.search_frame = tk.Frame(self.canvas)
		self.search_frame.bind("<Configure>", self.reset_scrollregion)
		
		pn_header = tk.Label(self.search_frame, text="PN")
		pn_header.grid(row=0, column=0, padx=10, pady=10)
		
		nsn_header = tk.Label(self.search_frame, text="NSN")
		nsn_header.grid(row=0, column=1, padx=10, pady=10)
		
		qty_header = tk.Label(self.search_frame, text="QTY")
		qty_header.grid(row=0, column=2, padx=10, pady=10)
		
		if not self.pn_var.get() and not self.nsn_var.get():
			display_list = self.inventory_list
		else:
			display_list = []
			
			for part in self.inventory_list:
				if self.pn_var.get():
					if self.nsn_var.get():
						if self.pn_var.get() in part[1] and self.nsn_var.get() in part[2]:
							display_list.append(part)
					else:
						if self.pn_var.get() in part[1]:
							display_list.append(part)
				else:
					if self.nsn_var.get() in part[2]:
						display_list.append(part)
						
		for i in range(1, len(display_list)+1):
			part = display_list[i-1]
			
			pn_search_label = tk.Label(self.search_frame, text=part[1])
			pn_search_label.grid(row=i, column=0, padx=10, pady=10)
			
			nsn_search_label = tk.Label(self.search_frame, text=part[2])
			nsn_search_label.grid(row=i, column=1, padx=10, pady=10)
			
			qty_search_label = tk.Label(self.search_frame, text=part[3])
			qty_search_label.grid(row=i, column=2, padx=10, pady=10)
			
			edit_button = tk.Button(self.search_frame, text="Edit")
			edit_button.grid(row=i, column=3, padx=10, pady=10)
			
		self.canvas.create_window((4,4),window=self.search_frame,anchor="nw")
		
		self.canvas.update_idletasks()
		self.canvas.configure(scrollregion=self.canvas.bbox('all'), yscrollcommand=self.scroll.set)
		
		self.canvas.pack(side="left",fill="both",expand=True)
		self.scroll.pack(side="right",fill="y")
		
	def search_window(self):
		sw = tk.Toplevel()
		sw.geometry('400x300')
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
		
	def add_stock(self):
		
		self.inventory_list.append([len(self.inventory_list)+1, self.pn_var.get(), self.nsn_var.get(), self.qty_var.get()])
		
		self.pn_var.set('')
		self.nsn_var.set('')
		self.qty_var.set(0)
		
		messagebox.showinfo("Stock Added", "Stock information has been saved!")
		
	def save_changes(self):
		
		with open("inventory.txt","wb") as fp:
			pickle.dump(self.inventory_list, fp)
		
	def show_inventory(self):
		
		for part in self.inventory_list:
			print(part[0],part[1],part[2],part[3])
	
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
		
		submit_button = tk.Button(win, text="Submit", command=self.add_stock)
		submit_button.grid(row=3, column=1, padx=5, pady=5, sticky="w")
		
	def inventory_window(self):
		t = tk.Toplevel()
		t.geometry('250x200')
		t.title("Inventory Management")
		
		5930011475211
		
		def _delete_window():
			try:
				self.save_changes()
				t.destroy()
			except:
				pass
				
		t.protocol("WM_DELETE_WINDOW", _delete_window)
		
		self.pn_var.trace("w", self.update_view)
		self.nsn_var.trace("w", self.update_view)
		
		search_button = tk.Button(t, text="Search Inventory", command=self.search_window)
		search_button.grid(row=0, column=0, padx=10, pady=10)
		
		add_button = tk.Button(t, text="Add Stock", command=self.add_stock_window)
		add_button.grid(row=1, column=0, padx=10, pady=10)
		
		save_button = tk.Button(t, text="Save Changes", command=self.save_changes)
		save_button.grid(row=2, column=0, padx=10, pady=10)
		