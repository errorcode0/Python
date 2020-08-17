

#Script created by:............Shawn Jones
#Date Created:.................5/26/2020
#Last Modified by:.............Shawn Jones

#Description: Script checks the data carrier spreadsheet on the S: drive and emails
#When there is a pending request.


####### Imports
import pandas as pd
import smtplib
import xlrd
import re
from pandas import ExcelWriter
import numpy
import openpyxl
from openpyxl import load_workbook
import os 
import csv
import datetime
import time
from re import search
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date 



##### File location and name
filelocation = "//nhc/nas/Infosys/EpicCare Ambulatory/_DC_Request"
filename = "DC_Remastered.xlsx"

cwd = os.getcwd()
os.chdir(filelocation)
os.listdir('.')
today = date.today()
D1 = today.strftime('%m/%d/%Y')

###


DC_Team = "<a href='mailto:Kimberly.Davenport@nortonhealthcare.org;Rusty.Fletcher@nortonhealthcare.org;Marquerite.Chaffee@nortonhealthcare.org;Shawn.Jones@nortonhealthcare.org?subject=NEW DC Request Response &body= I will move these'>Respond to the DC team"


##### Panda read in and list convert
df = pd.read_excel(filename)
#FDate = pd.to_datetime(df['Future_DC_Date'], format='%m-%d-%Y')
Enviornment =  df[['Environments']][df['Request Status'] == 'New'].values.tolist()
F1 = df[['Future_DC_Date']][df['Request Status'] == 'Future'].to_string()
Ftest = df[['Future_DC_Date']][df['Request Status'] == 'Future'].values.tolist()
FD = df['Future_DC_Date'].to_string()
NR = df['Request Status'].values.tolist()
Details = df[['Requestors Name','Package # POC to TST','Environments']][df['Request Status'] == 'New'].values.tolist()
Name =  df[['Requestors Name']][df['Request Status'] == 'New'].values.tolist()
PKGnum =  df[['Package # POC to TST']][df['Request Status'] == 'New'].values.tolist()
Enviornment =  df[['Environments']][df['Request Status'] == 'New'].values.tolist()
F2 = df[['Future_DC_Date']][df['Request Status'] == 'Future'].values.tolist()
F3 = df[['Requestors Name']][df['Request Status'] == 'Future'].values.tolist()
PKGnum2 =  df[['Package # POC to TST']][df['Request Status'] == 'Future'].values.tolist()

def Sendmail ():
    msg = MIMEMultipart()
    msg['from'] = "TestDC_Request@Nortonhealthcare.org"
    msg['To'] = "Shawn.Jones@nortonhealthcare.org"
    msg['Subject'] = "Test DC Request"
    #body = "<Style> body, h2{font-family: 'Raleway', Arial, sans-serif}" + "<p>"+ "<center>"   +"<h2 style='border:2px solid DodgerBlue;'>"+"There are pending Data Courier requests:" + "</h2>" + "</p>" + "</center>" +"<p>"+"Requesters Name | Pkg # | Enviornment" + "<br>" + str(Email).replace("[","").replace("]","<BR>").replace("'","").replace(",","").replace("'","").replace(".0","" ) + "<a href=file://nhc/nas/Infosys/EpicCare%20Ambulatory/_DC_Request/Data_Courier_Request.xlsx> Click to Open DC Spreadsheet"+ "</a>"+ "<br>"  + "<br>" + str(DC_Team) +"</a>" + "</style>"
    body = """<style> body, h2{font-family: 'Raleway', Arial, sans-serif}  
    
    
    </style>
     
       
    
 <table>
 <tr>
    <body>
    <h2 style='border:2px solid lawngreen;Background:lawngreen'>Pending DC Requests:</h2></p> 
    <table>
    <table cellspacing ='20'>
    <p>
    <tr>
        <td><b><u>Requesters Name</b></u><br>
        """ + str(Name).replace("[","").replace("]","<BR>").replace("'","").replace(",","").replace("'","").replace(".0","" ) + "</td>" +  """
        <td><b><u>Package Number</b></u><br>         
        """ + str(PKGnum).replace("[","").replace("]","<BR>").replace("'","").replace(",","").replace("'","").replace(".0","" )  + "</td>" + """
        <td><b><u>Enviornment</b></u><br>
        """ +  str(Enviornment).replace("[","").replace("]","<BR>").replace("'","").replace(",","").replace("'","").replace(".0","" ).replace("/\n/g","" )  +"</td>" + """
    </tr>
    </table>
    </p>
    <h2 style='border:2px solid darkorange;background:darkorange'>Future DC Request:</h2></p> 
    <p>
    
    <table>
    <table cellspacing ='20'>
    <tr>
        <td><b><u>Requesters Name</b></u><br>
        """ + str(F3).replace("[","").replace("]","<BR>").replace("'","").replace(",","").replace("'","").replace(".0","" ) + "</td>" +  """
        <td><b><u>Package Number</b></u><br>         
        """ + str(PKGnum2).replace("[","").replace("]","<BR>").replace("'","").replace(",","").replace("'","").replace(".0","" )  + "</td>" + """
        <td><b><u>Future Date</b></u><br>         
        """ + str(F2).replace("[","").replace("]","<BR>").replace("'","").replace(",","").replace("'","").replace(".0","" )  + "</td>" + """
    </tr>  
    </table>
    </p>
    <h2 style='border:2px solid DodgerBlue;Background:DodgerBlue'>HelpFul Links</h2></p>
    <p>
    <table>
    <td><b><u>
    <a href=file://nhc/nas/Infosys/EpicCare%20Ambulatory/_DC_Request/Data_Courier_Request.xlsx> Click to Open DC Spreadsheet </b></u><br>
    </table>    
    </p>
    <br>
    <p>
    <table>
    <td><b><u>""" + str(DC_Team) + """</b></u><br>
    </p>
    </td>
    </table>    
    </body>
  </tr>
  </table>
    
    """    
    
    msg.attach(MIMEText(body,'html'))
    print (msg)
    server = smtplib.SMTP("mail.nortonhealthcare.org", 25)
    server.sendmail(msg['from'], msg['To'],msg.as_string())
    server.quit()

    server = smtplib.SMTP("mail.nortonhealthcare.org", 25)
    server.send_message(msg)
    server.quit()
    
    
if 'New' in NR:
     Sendmail() 
   


     
     
print(NR)



#quit()




