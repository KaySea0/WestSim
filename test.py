# https://www.lifewire.com/what-are-the-outlook-com-smtp-server-settings-1170671
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

# wb = openpyxl.load_workbook('CRHT_0228100406 (1).xlsx')
# ws = wb.active
# max_row = ws.max_row

s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
s.starttls()
s.login(MY_ADDRESS,MY_PASSWORD)

msg = MIMEMultipart('related')

signature = """<p> Please provide a quote for the following: <br> <br>\
${PART_INFO} <br> 
Best, <br> </p>
<p style="color: blue;">Kyle Cook <br>
WestSim Engineering, Inc. <br>
7061 Grand National Dr. Suite 107A <br>
Orlando, FL 32819 <br></p>
<p> info@westsiminc.com <br>
Phone: (407) 963-6699 <br>
Fax: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(407) 203-4264 <br>
http://www.westsiminc.com/ <br>
<img src="cid:logo"> <br>
</p>
"""

msg['From'] = MY_ADDRESS
msg['To'] = "k.cook2499@gmail.com"
msg['Subject'] = "Test Email from Python"

msgBody = MIMEMultipart('alternative')
msg.attach(msgBody)

subMessage = Template(signature).substitute(PART_INFO = "P/N: 132456-789 - QTY 2 <br>")

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


# generic tkinter initialization
# root = tk.Tk()

# frame = tk.Frame(root)
# frame.pack()

# test_button = tk.Button(frame,text="Create Test Spreadsheet!",fg="red",command = create_sheet)
# test_button.pack(side=tk.LEFT)

# another_button = tk.Button(frame,text="Goodbye!",command=quit)
# another_button.pack(side=tk.LEFT)

# root.mainloop()