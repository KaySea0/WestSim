# https://stackoverflow.com/questions/15306631/how-do-i-create-child-windows-with-python-tkinter
# https://medium.freecodecamp.org/send-emails-using-code-4fcea9df63f - how to send emails automatically
import tkinter as tk
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from string import Template

MY_ADDRESS = "info@westsiminc.com"
MY_PASSWORD = "Sukkur%%1798"

def read_template(filename):
	with open(filename, 'r',) as template_file:
		template_content = template_file.read()
	return Template(template_content)

# wb = openpyxl.load_workbook('CRHT_0228100406 (1).xlsx')
# ws = wb.active
# max_row = ws.max_row

def send_emails():
	s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
	s.starttls()
	s.login(MY_ADDRESS,MY_PASSWORD)

	msg = MIMEMultipart('related')

	msg['From'] = MY_ADDRESS
	msg['To'] = "k.cook2499@gmail.com"
	msg['Subject'] = "Test Quote"

	msgBody = MIMEMultipart()
	msg.attach(msgBody)

	email_body = read_template('base_email.txt')
	subMessage = email_body.substitute(PART_INFO = "P/N: 132456-789 <br> QTY: 2 <br>")

	msgBody.attach(MIMEText(subMessage,'html'))

	fp = open('logo.png','rb')
	img = MIMEImage(fp.read())
	fp.close()
	img.add_header('Content-ID', '<logo>')
	msg.attach(img)

	s.send_message(msg)
	del msg

	s.quit()

# how to open / edit / save workbooks
# wb = openpyxl.load_workbook('CRHT_0228100406 (1).xlsx')
# ws = wb.active
# max_row = ws.max_row
# for i in range(1,max_row+1):
	# print(ws.cell(None,i,23).value)

# wb.save('Ship-Invoice-02142019.xlsx')

def create_window():
	t = tk.TopLevel(


# generic tkinter initialization
def init_window():
	root = tk.Tk()
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
	canvas.create_image(0,0, anchor="nw", image=img)
	
	test_button = tk.Button(frame,text="Create Test Spreadsheet!",fg="red",command = quit)
	test_button.grid(row=1,column=0,padx=10,pady=10)

	another_button = tk.Button(frame,text="Goodbye!",command=quit)
	another_button.grid(row=1,column=1,padx=10,pady=10)

	root.mainloop()
	
init_window()