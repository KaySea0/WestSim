# https://stackoverflow.com/questions/7546050/switch-between-two-frames-in-tkinter/49325719#49325719
import tkinter as tk
import os
import openpyxl

class WS_BidSheet(object):
	
	def __init__(self):
		self.sheet_names = os.listdir('Bid_Sheets')
		self.frame = None
		self.canvas = None
		self.sheet_name = tk.StringVar()
		self.scroll = None
	
	def show_info(self,row):
		pass
		
	def change_view(self, *args):
		if self.frame is not None:
			self.frame.destroy()
			
		self.frame = tk.Frame(self.canvas, background="#FFFFFF")
		
		cur_wb = openpyxl.load_workbook('Bid_Sheets/' + self.sheet_name.get())
		cur_ws = cur_wb.active
		
		for i in range(1,cur_ws.max_row):
			company_label = tk.Label(self.frame,text=str(cur_ws['A'+str(i)].value),bg="white")
			company_label.grid(row=i-1,column=0,sticky="nw",padx=5)
			
			pn_label = tk.Label(self.frame,text=str(cur_ws['B'+str(i)].value),bg="white")
			pn_label.grid(row=i-1,column=1,sticky="w",padx=5)
			
			qty_label = tk.Label(self.frame,text=str(cur_ws['C'+str(i)].value),bg="white")
			qty_label.grid(row=i-1,column=2,sticky="w",padx=5)
			
			info_button = tk.Button(self.frame,text="Add/View Info",command=lambda: self.show_info(i))
			info_button.grid(row=i-1,column=3,stick="e",padx=5)
			
		self.canvas.create_window((4,4),window=self.frame,anchor="nw")
		
		self.canvas.update_idletasks()
		self.canvas.configure(scrollregion=self.canvas.bbox('all'), yscrollcommand=self.scroll.set)
		
		self.canvas.pack(side="left",fill="both",expand=True)
		self.scroll.pack(side="right",fill="y")
		
	
	def bid_sheet_window(self):
		
		t = tk.Toplevel()
		t.geometry('450x300')
		t.title('Westsim Bid Sheets')
		self.sheet_names.reverse()
		
		self.sheet_name.set(self.sheet_names[0])
		self.sheet_name.trace('w', self.change_view)
		
		sheet_menu = tk.OptionMenu(t,self.sheet_name,*self.sheet_names)
		sheet_menu.pack(anchor="n",fill=tk.X)
		
		self.canvas = tk.Canvas(t,borderwidth=0,background="#FFFFFF")
		self.scroll = tk.Scrollbar(t,orient="vertical",command=self.canvas.yview)
	
		self.change_view()
		
		