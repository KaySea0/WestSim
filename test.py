# https://www.lifewire.com/what-are-the-outlook-com-smtp-server-settings-1170671
# https://medium.freecodecamp.org/send-emails-using-code-4fcea9df63f - how to send emails automatically
import tkinter as tk
import openpyxl
# import smtplib
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText

# MY_ADDRESS = "info@westsiminc.com"
# MY_PASSWORD = ""

# s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
# s.starttls()
# s.login(MY_ADDRESS,MY_PASSWORD)

# msg = MIMEMultipart()

# message = "This is a test email. Lets see if we can send a response to this!"

# msg['From'] = MY_ADDRESS
# msg['To'] = "k.cook2499@gmail.com"
# msg['Subject'] = "Test Email from Python"

# msg.attach(MIMEText(message,'plain'))

# s.send_message(msg)
# del msg

# s.quit()




# how to open / edit / save workbooks
wb = openpyxl.load_workbook('Ship-Invoice-02142019.xlsx')
ws = wb["DLAORDERS"]
print(ws.max_row)
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