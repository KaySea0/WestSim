# https://stackoverflow.com/questions/7546050/switch-between-two-frames-in-tkinter/49325719#49325719
# https://stackoverflow.com/questions/11980812/how-do-you-create-a-button-on-a-tkinter-canvas
import tkinter as tk
import os
import openpyxl

class WS_BidSheet(object):
	
	def __init__(self):
		self.sheet_names = os.listdir('Bid_Sheets')
		self.frame = None
		
	def change_view(self):
		if self.frame is not None:
			self.frame.destroy()
		
		# create canvas view for each cheat sheet
		
	def bid_sheet_window(self):
		
		t = tk.Toplevel()
		t.geometry('400x300')
		self.sheet_names.reverse()
		
		working_sheet = tk.StringVar(t)
		working_sheet.trace('w', self.change_view)
		working_sheet.set(self.sheet_names[0])
		
		sheet_menu = tk.OptionMenu(t,working_sheet,*self.sheet_names)
		sheet_menu.pack(expand=1,anchor="n",fill=tk.X)
		
		# bid_canvas = tk.Canvas(t,
		