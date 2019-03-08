# https://stackoverflow.com/questions/15306631/how-do-i-create-child-windows-with-python-tkinter
# https://medium.freecodecamp.org/send-emails-using-code-4fcea9df63f - how to send emails automatically
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import openpyxl
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from string import Template

# price break = p
# group email = g

# company name -> P/N -> QTY -> Info
# have way to add entry not from email
# days count down
# < $10k, 3 days
# > $10k, 2 weeks
# create lists based on files in Cheat_Sheet folder

# how to open / edit / save workbooks
# wb.save('Ship-Invoice-02142019.xlsx')

MY_ADDRESS = "info@westsiminc.com"
MY_PASSWORD = "Sukkur%%1798"

# MY_ADDRESS = "ali.kalwar@westsiminc.com"
# MY_PASSWORD = "Sukkur$$88"

class Frames(object):

	def __init__(self):
		self.ali_sheet_name = tk.StringVar()
		self.ali_sheet_name.set("default value")
		self.email_list = []
		
	def read_template(self,filename):
		with open(filename, 'r',) as template_file:
			template_content = template_file.read()
		return Template(template_content)	
		
	def open_file(self,file_destination):
		file_destination.set(filedialog.askopenfilename())
		
	def send_emails(self):
		s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
		s.starttls()
		s.login(MY_ADDRESS,MY_PASSWORD)
		
		for msg in self.email_list:
			s.send_message(msg)
		
	def process_ali_sheet(self, prev_window):
	
		ali_wb = openpyxl.load_workbook(self.ali_sheet_name.get())
		ali_ws = ali_wb.active
		ali_max_row = ali_ws.max_row
		
		with open('cage_dict.json') as f:
			cage_dict = json.load(f)
	
		message_list = []
		part_info = ""
		past_cage_code = ""
		email_body = self.read_template('base_email.txt')
		
		# to skip vendor, put contact info into second column of email
		
		if ali_max_row == 2 and cage_dict.get(str(ali_ws['V2'].value),"0") != "0":
			part_info = "P/N: " + str(ali_ws['X2'].value) + "<br> QTY: " + str(ali_ws['I2'].value) 
			
			if 'p' in cage_dict[str(ali_ws['V2'].value)]['options']:
				part_info += " or next price break"
				
			part_info += "<br><br>"
			
			sub_message = email_body.substitute(PART_INFO = part_info)
			to_address = cage_dict[str(ali_ws['V2'].value)]['email']
			message_list.append((sub_message,to_address))
			
		else:
			for i in range(2,ali_max_row+1):
				cur_cage_code = str(ali_ws['V' + str(i)].value)
				cur_PN = str(ali_ws['X' + str(i)].value)
				cur_QTY = str(ali_ws['I' + str(i)].value)
				
				if i == 2 and cage_dict.get(cur_cage_code,"0") != "0":
					part_info += "P/N: " + cur_PN + "<br> QTY: " + cur_QTY
					
					if 'p' in cage_dict[cur_cage_code]['options']:
						part_info += " or next price break"
						
					part_info += "<br><br>"
					past_cage_code = cur_cage_code
					
				elif cage_dict.get(cur_cage_code,"0") != "0":
					if cur_cage_code == past_cage_code and 'g' in cage_dict[cur_cage_code]['options']:
						part_info += "P/N: " + cur_PN + "<br> QTY: " + cur_QTY
						
						if 'p' in cage_dict[cur_cage_code]['options']:
							part_info += " or next price break"
							
						part_info += "<br><br>"
						
					else:
						if cage_dict.get(past_cage_code,"0") != "0":
							sub_message = email_body.substitute(PART_INFO = part_info)
							to_address = cage_dict[past_cage_code]['email']
							message_list.append((sub_message,to_address))
						
						past_cage_code = cur_cage_code
						part_info = "P/N: " + cur_PN + "<br> QTY: " + cur_QTY
						
						if 'p' in cage_dict[cur_cage_code]['options']:
							part_info += " or next price break"
						
						part_info += "<br><br>"
						
					if i == ali_max_row:
						sub_message = email_body.substitute(PART_INFO = part_info)
						to_address = cage_dict[past_cage_code]['email']
						message_list.append((sub_message,to_address))
						
						part_info = "P/N: " + cur_PN + "<br> QTY: " + cur_QTY
						
						if 'p' in cage_dict[cur_cage_code]['options']:
							part_info += " or next price break"
							
						part_info += "<br><br>"
						
						sub_message = email_body.substitute(PART_INFO = part_info)
						to_address = cage_dict[cur_cage_code]['email']
						message_list.append((sub_message,to_address))
						
						
				elif part_info != "":
					sub_message = email_body.substitute(PART_INFO = part_info)
					to_address = cage_dict[past_cage_code]['email']
					message_list.append((sub_message,to_address))
					
					part_info = ""
					past_cage_code = cur_cage_code
		
		count = {'value': 0}
	
		prev_window.destroy()
		t = tk.Toplevel()
		t.title("Email Confirmation")
		t.geometry('500x600')
		t.rowconfigure(0,weight=1)
		t.columnconfigure(0,weight=1)
		
		email_frame = tk.Frame(t)
		email_frame.grid_rowconfigure(0,weight=1)
		email_frame.grid_columnconfigure(0,weight=1)
		email_frame.grid(row=0,column=0)
		
		scroll = tk.Scrollbar(email_frame, orient="vertical")
		scroll.grid(row=0,column=1,sticky="ns")
		
		email_preview = tk.Text(email_frame,height=500,width=75,yscrollcommand=scroll.set)
		email_preview.insert("1.0","TO: " + message_list[count['value']][1] + "\n" + message_list[count['value']][0].replace("<br><br>","\n").replace("<br>","\n"))
		email_preview.grid(row=0,column=0,sticky="nwe",padx=10,pady=10)
		
		scroll.config(command=email_preview.yview)
		
		confirmation_label = tk.Label(t,text="Does this email look correct?")
		confirmation_label.grid(row=1,column=0,sticky="w",padx=5)
		
		def confirm_email():
			msg = MIMEMultipart('related')

			msg['From'] = MY_ADDRESS
			# msg['To'] = message_list[count['value']][1]
			msg['To'] = "k.cook2499@gmail.com"
			msg['Subject'] = "Quote"

			msgBody = MIMEMultipart()
			msg.attach(msgBody)

			msgBody.attach(MIMEText("TO: " + message_list[count['value']][1] + "<br>" + message_list[count['value']][0],'html'))

			fp = open('logo.png','rb')
			img = MIMEImage(fp.read())
			fp.close()
			img.add_header('Content-ID', '<logo>')
			msg.attach(img)
			
			self.email_list.append(msg)
			
			count['value'] += 1
			
			if count['value'] != len(message_list):
				email_preview.delete('1.0',tk.END)
				email_preview.insert("1.0","TO: " + message_list[count['value']][1] + "\n" + message_list[count['value']][0].replace("<br><br>","\n").replace("<br>","\n"))
			else:
				self.send_emails()
				t.destroy()
				tk.tkMessageBox.showinfo("Email Confirmation", "All " + str(len(self.email_list)) + " emails have been sent!")
		
		confirm_button = tk.Button(t,text="Yes",command = confirm_email)
		confirm_button.grid(row=2,column=0,stick="w",padx=5)
		
	def main_frame(self,root):
		root.title("Westsim Engineering")
		root.geometry('500x300')
		
		frame = tk.Frame(root)
		frame.config(bg="white")
		
		rows = 0
		while rows < 50:
			frame.rowconfigure(rows,weight=1)
			frame.columnconfigure(rows,weight=1)
			rows += 1
			
		frame.pack(side=tk.LEFT,anchor="nw")
		
		canvas = tk.Canvas(frame,width=99,height=39)
		canvas.grid(row=0,column=0)
		img = tk.PhotoImage(file="logo.gif")
		canvas.image = img
		canvas.create_image(0,0, anchor="nw", image=img)
		
		test_button = tk.Button(frame,text="Open a new window!",fg="blue",command = self.sub_window)
		test_button.grid(row=1,column=0,padx=10,pady=10)

		another_button = tk.Button(frame,text="Goodbye!",command=quit)
		another_button.grid(row=1,column=1,padx=10,pady=10)
		
		alert_test = tk.Button(frame,text="Let's open an alert window!",command=lambda: tk.messagebox.showinfo(("Email Confirmation", "This is a test alert!")))
		
	def sub_window(self):
	
		t = tk.Toplevel()
		t.title("Sub-window")
		t.geometry('400x200')
		
		rows = 0
		while rows < 50:
			t.rowconfigure(rows,weight=1)
			t.columnconfigure(rows,weight=1)
			rows += 1
		
		intro_label = tk.Label(t, text="Select ALICORP spreadsheet you wish to process:")
		intro_label.grid(row=0,column=0,padx=5,pady=10)
		
		sheet_name_text = tk.Entry(t, state="disabled", textvariable=self.ali_sheet_name, width=50)
		sheet_name_text.grid(row=1,column=0,padx=10,pady=10,sticky="w")
		
		process_button = tk.Button(t, text="Process Spreadsheet", state="disabled", command=lambda: self.process_ali_sheet(t))
		process_button.grid(row=2,column=0,padx=5)
		
		def browse_function():
			self.open_file(self.ali_sheet_name)
			process_button.configure(state="normal")
		
		browse_button = tk.Button(t, text="Browse", command = browse_function)
		browse_button.grid(row=1,column=1,padx=5,sticky="w")
		

	
root = tk.Tk()
app = Frames()
app.main_frame(root)
root.mainloop()