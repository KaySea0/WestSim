import tkinter as tk
import os
import openpyxl

class WS_BidSheet(object):
	
	def __init__(self):
	
		self.bid_folder = None # location of bid sheet directory
		self.sheet_names = None # list of all bid sheets in folder
		self.frame = None # bid sheet frame that switches as different dates are selected
		self.canvas = None # main canvas that houses bid sheet frame
		self.scroll = None # scrollbar that controls canvas
		self.cur_wb = None # workbook of currently selected bid sheet
		self.sheet_name = tk.StringVar() # name of currently selected bid sheet
	
	# # #
	# Method: update_list
	# Input: n/a
	# Utility: 
	#   Update list of all bid sheets in specified folder
	# # #
	def update_list(self, folder):
		self.bid_folder = folder
		self.sheet_names = os.listdir(self.bid_folder)
		# del self.sheet_names[-1]
		
	# # #
	# Method: submit_info
	# Input:
	#   row - row of current bid sheet that is being modified in notebook portion
	#   text - notes that is grabbed from spreadsheet / edited by user
	#   window - notepad window
	# Utility:
	#   Take text from notepad portion, update & save bid sheet, and destroy window
	# # #
	def submit_info(self,row,text,window):
		
		# get current sheet and save data to correct cell
		cur_ws = self.cur_wb.active
		cur_ws['D'+str(row)] = text
		self.cur_wb.save(self.bid_folder + "/" + self.sheet_name.get())
		
		# close notepad window
		window.destroy()
	
	# # #
	# Method: show_info
	# Input:
	#   row - row of current bid sheet to display info of
	# Utility:
	#   Open 'notepad' window that displays info of selected contract to aid in bidding. Allows user to edit info as quotes are submitted.
	# # #
	def show_info(self,row):
	
		cur_ws = self.cur_wb.active # get current bid sheet
		
		# if there is info for the certain selection, pull it for viewing
		if cur_ws['D'+str(row)].value:
			info = cur_ws['D'+str(row)].value
		else:
			info = ""
		
		# create 'notepad' window for viewing extra data about potential contracts
		info_window = tk.Toplevel()
		info_window.geometry('300x400')
		info_window.title("Bid Info")
		info_window.rowconfigure(0,weight=1)
		info_window.columnconfigure(0,weight=1)
		
		# frame that contains notepad text widget and scrollbar
		info_frame = tk.Frame(info_window)
		info_frame.grid_rowconfigure(0,weight=1)
		info_frame.grid_columnconfigure(0,weight=1)
		info_frame.grid(row=0,column=0)
		
		# scrollbar to control text widget
		scroll = tk.Scrollbar(info_frame, orient="vertical")
		scroll.grid(row=0,column=1,sticky="ns")
		
		# text widget for notepad info
		info_text = tk.Text(info_frame,height=200,width=30,yscrollcommand=scroll.set)
		info_text.insert("1.0",info)
		info_text.grid(row=0,column=0,sticky="nwe",padx=10,pady=10)
		
		scroll.config(command=info_text.yview)
		
		# frame for all other UI elements
		user_frame = tk.Frame(info_window)
		user_frame.grid(row=1,column=0,sticky="sw",padx=5,pady=5)
		
		submit_button = tk.Button(user_frame,text="Save/Close Window",command=lambda: self.submit_info(row,info_text.get("1.0",tk.END),info_window))
		submit_button.grid(row=1,column=0,sticky="n",padx=80)
	
	# # #
	# Method: change_view
	# Input: n/a
	# Utility:
	#   Delete prior window (if one exists) and populate list of entries in bid sheet based on what is selected in option menu. 
	# # #
	def change_view(self, *args):
	
		# if a frame is currently being displayed, destroy it to make room for new frame
		if self.frame is not None:
			self.frame.destroy()
			
		self.frame = tk.Frame(self.canvas, background="#FFFFFF")
		
		# open up selected bid sheet
		self.cur_wb = openpyxl.load_workbook(self.bid_folder + "/" + self.sheet_name.get())
		cur_ws = self.cur_wb.active
		
		# headers for entry list
		tk.Label(self.frame,text="Company Name",bg="white",font="bold").grid(row=0,column=0,sticky="nw",padx=5)
		tk.Label(self.frame,text="P/N",bg="white",font="bold").grid(row=0,column=1,sticky="w",padx=5)
		tk.Label(self.frame,text="QTY",bg="white",font="bold").grid(row=0,column=2,sticky="w",padx=5)
		
		# populate each row of bid sheet with...
		for i in range(1,cur_ws.max_row+1):
		
			# ...vendor name...
			company_label = tk.Label(self.frame,text=cur_ws['A'+str(i)].value,bg="white")
			company_label.grid(row=i,column=0,sticky="nw",padx=5)
			
			# ...part number...
			pn_label = tk.Label(self.frame,text=str(cur_ws['B'+str(i)].value),bg="white")
			pn_label.grid(row=i,column=1,sticky="w",padx=5)
			
			# ...quantity...
			qty_label = tk.Label(self.frame,text=str(cur_ws['C'+str(i)].value),bg="white")
			qty_label.grid(row=i,column=2,sticky="w",padx=5)
			
			# ...and a button to show/edit extra info for bidding purposes 
			info_button = tk.Button(self.frame,text="Add/View Info",command=lambda i=i: self.show_info(i))
			info_button.grid(row=i,column=3,sticky="e",padx=5)
		
		# once frame is populated, add to main canvas
		self.canvas.create_window((4,4),window=self.frame,anchor="nw")
		
		self.canvas.update_idletasks()
		self.canvas.configure(scrollregion=self.canvas.bbox('all'), yscrollcommand=self.scroll.set)
		
		self.canvas.pack(side="left",fill="both",expand=True)
		self.scroll.pack(side="right",fill="y")
		
	# # #
	# Method: bid_sheet_window
	# Input: n/a
	# Utility:
	#   Create initial referece objects for 'bid sheet' window and call for a view change based on default sheet selection
	# # #
	def bid_sheet_window(self):
		
		# create main 'bid sheet' window
		t = tk.Toplevel()
		t.geometry('500x300')
		t.title('Westsim Bid Sheets')
		
		# make it so most recent bid sheet is shown by default
		self.sheet_names.reverse()
		
		# set initial value to most recent sheet and set tracer function to change_view
		self.sheet_name.set(self.sheet_names[0])
		self.sheet_name.trace('w', self.change_view)
		
		# option menu object that allows user to select which sheet they want to look at
		sheet_menu = tk.OptionMenu(t,self.sheet_name,*self.sheet_names)
		sheet_menu.pack(anchor="n",fill=tk.X)
		
		self.canvas = tk.Canvas(t,borderwidth=0,background="#FFFFFF")
		self.scroll = tk.Scrollbar(t,orient="vertical",command=self.canvas.yview)
	
		self.change_view()
		
		