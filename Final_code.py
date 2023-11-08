# -*- coding: utf-8 -*-
"""
Created on Fri Nov  3 14:00:08 2023

@author: LUMBINI ELITE
"""

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st



with st.form("WEEKLY MAILING"):
        st.write("Please select the user group:")
        input_type = st.selectbox("Input Type", ["PR","PO"])
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
        submitted = st.form_submit_button("Submit")



def prmail():    
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(uploaded_file)
    
    unique_emails = df['Pr Creator Email Id'].unique()
    
    # Set up your email server and credentials
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_username = "nitheshshetty32@gmail.com"
    smtp_password = "cklt lhbg spqa ihbr"
    
    for email in unique_emails:
        # Filter rows for the current email
        recipient_rows = df[df['Pr Creator Email Id'] == email]
    
        # Prepare the email content as an HTML table
        subject = "PR DETAILS"
        from_email = "nitheshshetty32@gmail.com"
        to_email = email  # The recipient's email address
    
        # Create an HTML message object
        msg = MIMEMultipart()
        msg["From"] = from_email
        msg["To"] = to_email
        msg["Subject"] = subject
    
        # Create the email body as an HTML table
        email_body = """
        <html>
        <body>
        <h2>PR Details</h2>
        <table border="1">
        <tr>
        <th>Sn.No</th>
        <th>Company Code</th>
        <th>Vendor Number</th>
        <th>Vendor Name</th>
        <th>Reconciliation GL</th>
        <th>GL Account Description</th>
        <th>Accounting Document Num</th>
        <th>Fiscal Year</th>
        <th>Line item</th>
        <th>Special GL Indicator</th>
        <th>Special GL Indicator Description</th>
        <th>Posting Date</th>
        <th>Document Date</th>
        <th>Net Due Date</th>
        <th>Reference</th>
        <th>Document Type</th>
        <th>Posting Key</th>
        <th>Debit Credit Indicator</th>
        <th>Item Text</th>
        <th>Profit Center</th>
        <th>Minority Indicator</th>
        <th>Purchasing Document</th>
        <th>Purchasing Group</th>
        <th>PR Creator ID</th>
        <th>PR Creator Name</th>
        <th>Pr Creator Email Id</th>
        <th>PO Creator ID</th>
        <th>PO Creator Name</th>
        <th>Po Creator Email Id</th>
        <th>Document Header Text</th>
        <th>Currency</th>
        <th>Ageing</th>
        <th>Amount Document Currency</th>
        <th>Amount Local Currency</th>
        <th>No Due Date LC</th>
        <th>01-07 Ageing Amount</th>
        <th>08-15 Ageing Amount</th>
        <th>16-30 Ageing Amount</th>
        <th>31-60 Ageing Amount</th>
        <th>61-90 Ageing Amount</th>
        <th>91-120 Ageing Amount</th>
        <th>121-180 Ageing Amountr</th>
        <th>181-365 Ageing Amount</th>
        <th>More than 365 Ageing Amount</th>
        <table style="table-layout: auto; width: 100%;">
        </tr>
        """
    
        for index, row in recipient_rows.iterrows():
            # Modify this part according to your column names
            email_body += f"<tr><td>{row['Sn.No']}</td><td>{row['Company Code']}</td><td>{row['Vendor Number']}</td><td>{row['Reconciliation GL']}</td><td>{row['GL Account Description']}</td><td>{row['Accounting Document Num']}</td><td>{row['Fiscal Year']}</td><td>{row['Line item']}</td><td>{row['Special GL Indicator']}</td><td>{row['Special GL Indicator Description']}</td><td>{row['Posting Date']}</td><td>{row['Document Date']}</td><td>{row['Net Due Date']}</td><td>{row['Reference']}</td><td>{row['Document Type']}</td><td>{row['Posting Key']}</td><td>{row['Debit Credit Indicator']}</td><td>{row['Item Text']}</td><td>{row['Profit Center']}</td><td>{row['Minority Indicator']}</td><td>{row['Purchasing Document']}</td><td>{row['Purchasing Group']}</td><td>{row['PR Creator ID']}</td><td>{row['PR Creator Name']}</td><td>{row['Pr Creator Email Id']}</td><td>{row['PO Creator ID']}</td><td>{row['PO Creator Name']}</td><td>{row['Po Creator Email Id']}</td><td>{row['Document Header Text']}</td><td>{row['Currency']}</td><td>{row['Ageing']}</td><td>{row['Amount Document Currency']}</td><td>{row['Amount Local Currency']}</td><td>{row['No Due Date LC']}</td><td>{row['01-07 Ageing Amount']}</td><td>{row['08-15 Ageing Amount']}</td><td>{row['16-30 Ageing Amount']}</td><td>{row['31-60 Ageing Amount']}</td><td>{row['61-90 Ageing Amount']}</td><td>{row['91-120 Ageing Amount']}</td><td>{row['121-180 Ageing Amount']}</td><td>{row['181-365 Ageing Amount']}</td><td>{row['More than 365 Ageing Amount']}</td></tr>"
    
        email_body += """
        </table>
        </body>
        </html>
        """
    
        msg.attach(MIMEText(email_body, "html"))
    
        # Connect to the SMTP server and send the email
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)
            text = msg.as_string()
            server.sendmail(from_email, to_email, text)
            server.quit()
            st.write("Email sent to ",to_email)
        except Exception as e:
            st.write(f"Error sending email to",to_email)
def pomail():
     
     # Read the Excel file into a Pandas DataFrame
     df = pd.read_excel(uploaded_file)
     
     unique_emails = df['Po Creator Email Id'].unique()
     
     # Set up your email server and credentials
     smtp_server = "smtp.gmail.com"
     smtp_port = 587
     smtp_username = "nitheshshetty32@gmail.com"
     smtp_password = "cklt lhbg spqa ihbr"
     
     for email in unique_emails:
         # Filter rows for the current email
         recipient_rows = df[df['Po Creator Email Id'] == email]
     
         # Prepare the email content as an HTML table
         subject = "PO DETAILS"
         from_email = "nitheshshetty32@gmail.com"
         to_email = email  # The recipient's email address
     
         # Create an HTML message object
         msg = MIMEMultipart()
         msg["From"] = from_email
         msg["To"] = to_email
         msg["Subject"] = subject
     
         # Create the email body as an HTML table
         email_body = """
         <html>
         <body>
         <h2>PO Details</h2>
         <table border="1">
         <tr>
         <th>Sn.No</th>
        <th>Company Code</th>
        <th>Vendor Number</th>
        <th>Vendor Name</th>
        <th>Reconciliation GL</th>
        <th>GL Account Description</th>
        <th>Accounting Document Num</th>
        <th>Fiscal Year</th>
        <th>Line item</th>
        <th>Special GL Indicator</th>
        <th>Special GL Indicator Description</th>
        <th>Posting Date</th>
        <th>Document Date</th>
        <th>Net Due Date</th>
        <th>Reference</th>
        <th>Document Type</th>
        <th>Posting Key</th>
        <th>Debit Credit Indicator</th>
        <th>Item Text</th>
        <th>Profit Center</th>
        <th>Minority Indicator</th>
        <th>Purchasing Document</th>
        <th>Purchasing Group</th>
        <th>PR Creator ID</th>
        <th>PR Creator Name</th>
        <th>Pr Creator Email Id</th>
        <th>PO Creator ID</th>
        <th>PO Creator Name</th>
        <th>Po Creator Email Id</th>
        <th>Document Header Text</th>
        <th>Currency</th>
        <th>Ageing</th>
        <th>Amount Document Currency</th>
        <th>Amount Local Currency</th>
        <th>No Due Date LC</th>
        <th>01-07 Ageing Amount</th>
        <th>08-15 Ageing Amount</th>
        <th>16-30 Ageing Amount</th>
        <th>31-60 Ageing Amount</th>
        <th>61-90 Ageing Amount</th>
        <th>91-120 Ageing Amount</th>
        <th>121-180 Ageing Amountr</th>
        <th>181-365 Ageing Amount</th>
        <th>More than 365 Ageing Amount</th>
         </tr>
         """
     
         for index, row in recipient_rows.iterrows():
             # Modify this part according to your column names
             email_body += f"<tr><td>{row['Sn.No']}</td><td>{row['Company Code']}</td><td>{row['Vendor Number']}</td><td>{row['Reconciliation GL']}</td><td>{row['GL Account Description']}</td><td>{row['Accounting Document Num']}</td><td>{row['Fiscal Year']}</td><td>{row['Line item']}</td><td>{row['Special GL Indicator']}</td><td>{row['Special GL Indicator Description']}</td><td>{row['Posting Date']}</td><td>{row['Document Date']}</td><td>{row['Net Due Date']}</td><td>{row['Reference']}</td><td>{row['Document Type']}</td><td>{row['Posting Key']}</td><td>{row['Debit Credit Indicator']}</td><td>{row['Item Text']}</td><td>{row['Profit Center']}</td><td>{row['Minority Indicator']}</td><td>{row['Purchasing Document']}</td><td>{row['Purchasing Group']}</td><td>{row['PR Creator ID']}</td><td>{row['PR Creator Name']}</td><td>{row['Pr Creator Email Id']}</td><td>{row['PO Creator ID']}</td><td>{row['PO Creator Name']}</td><td>{row['Po Creator Email Id']}</td><td>{row['Document Header Text']}</td><td>{row['Currency']}</td><td>{row['Ageing']}</td><td>{row['Amount Document Currency']}</td><td>{row['Amount Local Currency']}</td><td>{row['No Due Date LC']}</td><td>{row['01-07 Ageing Amount']}</td><td>{row['08-15 Ageing Amount']}</td><td>{row['16-30 Ageing Amount']}</td><td>{row['31-60 Ageing Amount']}</td><td>{row['61-90 Ageing Amount']}</td><td>{row['91-120 Ageing Amount']}</td><td>{row['121-180 Ageing Amount']}</td><td>{row['181-365 Ageing Amount']}</td><td>{row['More than 365 Ageing Amount']}</td></tr>"
     
         email_body += """
         </table>
         </body>
         </html>
         """
     
         msg.attach(MIMEText(email_body, "html"))
     
         # Connect to the SMTP server and send the email
         try:
             server = smtplib.SMTP(smtp_server, smtp_port)
             server.starttls()
             server.login(smtp_username, smtp_password)
             text = msg.as_string()
             server.sendmail(from_email, to_email, text)
             server.quit()
             st.write("Email sent to ",to_email)
         except Exception as e:
             st.write(f"Error sending email to",to_email)

if input_type=="PR":
    prmail()
else:
    pomail()
