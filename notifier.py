import pandas as pd
import smtplib
import ssl
from datetime import datetime
from openpyxl import load_workbook
from config import username, password, smtp_server, smtp_port  # Import sensitive information

# Load the Excel file
wb = load_workbook(filename='Your_File.xlsx')
sheet = wb.active

# Read the Excel file into a pandas DataFrame
df = pd.read_excel('Your_File.xlsx')

# Get today's date
today = datetime.today().strftime('%Y-%m-%d')

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    # Check if the runout date is due
    if row['Runout-Date'] <= today:
        # Create a secure SSL context
        context = ssl.create_default_context()

        # Use the SMTP client
        with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
            server.login(username, password)

            # Send email
            subject = "Your runout date is due"
            body = "Dear {},\n\nYour runout date is due. Please take the necessary actions.".format(row['Name'])
            message = f'Subject: {subject}\n\n{body}'
            server.sendmail(username, row['Email'], message)
