import tkinter as tk
from WS_Email import WS_Email
from WS_BidSheet import WS_BidSheet
from WS_Config import WS_Config
from WS_Contract import WS_Contract
from WS_Inventory import WS_Inventory
from WS_Shipping import WS_Shipping
import json
import os
import sys
from myimages import *

class Westsim_App(object):
	
	def __init__(self):
		
		self.config_dict = None
		
		self.WS_Email = WS_Email()
		self.WS_BidSheet = WS_BidSheet()
		self.WS_Config = WS_Config()
		self.WS_Contract = WS_Contract()
		self.WS_Inventory = WS_Inventory()
		self.WS_Shipping = WS_Shipping()
		self.ws = 0
		self.hs = 0
		
	def start_bid_sheet(self):
		self.config_dict = json.load(open('config_dict.json'))
		self.WS_BidSheet.update_list(self.config_dict['bid'])
		self.WS_BidSheet.bid_sheet_window()
		
	def start(self,root):
		root.title("Westsim Engineering")
		
		self.ws = root.winfo_screenwidth()
		self.hs = root.winfo_screenheight()
		
		w = 300
		h = 190
		
		x = (self.ws/2) - (w/2)
		y = (self.hs/2) - (h/2)
		
		root.geometry('%dx%d+%d+%d' % (w,h,x,y))
		
		frame = tk.Frame(root)
		frame.config(bg="white")
		
		rows = 0
		while rows < 50:
			frame.rowconfigure(rows,weight=1)
			frame.columnconfigure(rows,weight=1)
			rows += 1
			
		frame.pack(side=tk.LEFT,anchor="nw")
		
		canvas = tk.Canvas(frame,width=99,height=39)
		canvas.grid(row=0,column=0, columnspan=2, padx=5, pady=5)
		img = tk.PhotoImage(data=logo_string)
		canvas.image = img
		canvas.create_image(0,0, anchor="nw", image=img)
		
		email_button = tk.Button(frame,text="Send Quote Emails",command= self.WS_Email.email_window)
		email_button.grid(row=1,column=0,padx=10,pady=10)
		
		bidsheet_button = tk.Button(frame,text="Open Bid Sheets",command= self.start_bid_sheet)
		bidsheet_button.grid(row=1,column=1,padx=10,pady=10)
		
		contract_button = tk.Button(frame,text="Contract Management", command= self.WS_Contract.contract_window)
		contract_button.grid(row=2,column=0,padx=10,pady=10)
		
		inventory_button = tk.Button(frame,text="Inventory Management", command= self.WS_Inventory.inventory_window)
		inventory_button.grid(row=2,column=1,padx=10,pady=10)
		
		shipping_button = tk.Button(frame,text="Shipping Management", command= self.WS_Shipping.shipping_window)
		shipping_button.grid(row=3,column=0,padx=10,pady=10)
		
		config_button = tk.Button(frame,text="Config",command = self.WS_Config.config_window)
		config_button.grid(row=3,column=1,padx=10,pady=10)

if __name__ == '__main__':		
	root = tk.Tk()
	app = Westsim_App()
	
	def _delete_window():
		try:
			app.WS_Contract.save_changes()
			app.WS_Inventory.save_changes()
			app.WS_Shipping.save_changes()
			root.destroy()
		except:
			root.destroy()
			
	root.protocol("WM_DELETE_WINDOW", _delete_window)
		
	app.start(root)
	root.mainloop()
	