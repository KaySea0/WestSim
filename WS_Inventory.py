import tkinter as tk
import json

class WS_Inventory(object):
	
	def __init__(self):
		
		try:
			self.inventory_list = json.load(open('inventory.json'))
		except:
			self.inventory_list = None
			
	def inventory_window(self):
		t = tk.Toplevel()
		t.geometry('250x400')