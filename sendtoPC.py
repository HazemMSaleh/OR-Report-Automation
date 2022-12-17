import pandas as pd
from datetime import datetime
import os
import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

#access OR report for buyer emails
dfemails = pd.read_excel('Z:\Compliance Reporting Team\Over Received ERS\Automation\Over Received Automation - Test 12-16.xlsm', sheet_name= 'Buyer List') 
buyer_names = dfemails["Buyer"].unique()

for name in buyer_names:
    email = dfemails[dfemails["Buyer"] == name]['Email'].values[0] #query buyer's name
    buyer_path = '"Z:\\Compliance Reporting Team\\Over Received ERS\\Automation\data\\' + name + '.xlsx"' #define location of buyer's sheet
    print(name)
    print(buyer_path)
    mail = outlook.CreateItem(0) 
    mail.Subject = '[Compliance Reporting Team Over Received PO Auto Output]  ' + datetime.now().strftime('%#d %b %Y %H:%M')
    mail.To = email
    mail.HTMLBody = """
    <html><body> 
    Hi """ + name + """,<br><br>
    Please find the Over Received items awaiting your update for this week <a href = """ + buyer_path + """>here</a>. NOTE: This document is on the share drive. You may edit the contents of the column titled ""Comments"", and may change the color of any cell on the worksheet, but please do not delete any columns.<br><br><i>
    (This is an automated message from the US PC Team Over Received PO Automation)
    </i></body>
    """
    mail.Display(True) #change to mail.Send if want to send email directly
    