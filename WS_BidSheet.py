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
		self.cur_wb = None
		
	def submit_info(self,row,text,window):
		cur_ws = self.cur_wb.active
		cur_ws['D'+str(row)] = text
		self.cur_wb.save('Bid_Sheets/' + self.sheet_name.get())
		window.destroy()
	
	def show_info(self,row):
		# pass
		cur_ws = self.cur_wb.active
		
		def save_info(text):
			cur_ws['D'+str(row)] = text
		
		if cur_ws['D'+str(row)].value:
			info = cur_ws['D'+str(row)].value
		else:
			info = ""
		
		info_window = tk.Toplevel()
		info_window.geometry('300x400')
		info_window.title("Bid Info")
		info_window.rowconfigure(0,weight=1)
		info_window.columnconfigure(0,weight=1)
		
		info_frame = tk.Frame(info_window)
		info_frame.grid_rowconfigure(0,weight=1)
		info_frame.grid_columnconfigure(0,weight=1)
		info_frame.grid(row=0,column=0)
		
		scroll = tk.Scrollbar(info_frame, orient="vertical")
		scroll.grid(row=0,column=1,sticky="ns")
		
		info_text = tk.Text(info_frame,height=200,width=30,yscrollcommand=scroll.set)
		info_text.insert("1.0",info)
		info_text.grid(row=0,column=0,sticky="nwe",padx=10,pady=10)
		
		scroll.config(command=info_text.yview)
		
		user_frame = tk.Frame(info_window)
		user_frame.grid(row=1,column=0,sticky="sw",padx=5,pady=5)
		
		submit_button = tk.Button(user_frame,text="Save/Close Window",command=lambda: self.submit_info(row,info_text.get("1.0",tk.END),info_window))
		submit_button.grid(row=1,column=0,sticky="w")
		
	def change_view(self, *args):
		if self.frame is not None:
			self.frame.destroy()
			
		self.frame = tk.Frame(self.canvas, background="#FFFFFF")
		
		self.cur_wb = openpyxl.load_workbook('Bid_Sheets/' + self.sheet_name.get())
		cur_ws = self.cur_wb.active
		
		tk.Label(self.frame,text="Company Name",bg="white",font="bold").grid(row=0,column=0,sticky="nw",padx=5)
		tk.Label(self.frame,text="P/N",bg="white",font="bold").grid(row=0,column=1,sticky="w",padx=5)
		tk.Label(self.frame,text="QTY",bg="white",font="bold").grid(row=0,column=2,sticky="w",padx=5)
		
		for i in range(1,cur_ws.max_row+1):
			company_label = tk.Label(self.frame,text=cur_ws['A'+str(i)].value,bg="white")
			company_label.grid(row=i,column=0,sticky="nw",padx=5)
			
			pn_label = tk.Label(self.frame,text=str(cur_ws['B'+str(i)].value),bg="white")
			pn_label.grid(row=i,column=1,sticky="w",padx=5)
			
			qty_label = tk.Label(self.frame,text=str(cur_ws['C'+str(i)].value),bg="white")
			qty_label.grid(row=i,column=2,sticky="w",padx=5)
			
			info_button = tk.Button(self.frame,text="Add/View Info",command=lambda i=i: self.show_info(i))
			info_button.grid(row=i,column=3,sticky="e",padx=5)
			
		self.canvas.create_window((4,4),window=self.frame,anchor="nw")
		
		self.canvas.update_idletasks()
		self.canvas.configure(scrollregion=self.canvas.bbox('all'), yscrollcommand=self.scroll.set)
		
		self.canvas.pack(side="left",fill="both",expand=True)
		self.scroll.pack(side="right",fill="y")
		
	
	def bid_sheet_window(self):
		
		t = tk.Toplevel()
		t.geometry('500x300')
		t.title('Westsim Bid Sheets')
		self.sheet_names.reverse()
		
		self.sheet_name.set(self.sheet_names[0])
		self.sheet_name.trace('w', self.change_view)
		
		sheet_menu = tk.OptionMenu(t,self.sheet_name,*self.sheet_names)
		sheet_menu.pack(anchor="n",fill=tk.X)
		
		self.canvas = tk.Canvas(t,borderwidth=0,background="#FFFFFF")
		self.scroll = tk.Scrollbar(t,orient="vertical",command=self.canvas.yview)
	
		self.change_view()
		
		