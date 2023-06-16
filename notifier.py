import pandas as pd
import smtplib
import ssl
from datetime import datetime
from openpyxl import load_workbook
from config import username, password, smtp_server, smtp_port, excel_file  # Import sensitive information

# Load the Excel file
wb = load_workbook(filename=excel_file)
sheet = wb.active

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(excel_file)

# Get today's date
today = datetime.today()

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    # Convert 'Runout-Date' to datetime if it's not already
    runout_date = pd.to_datetime(row['Runout-Date'])

    # Check if the runout date is due
    if runout_date.date() <= today.date():
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
