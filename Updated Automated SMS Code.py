import os
import pyodbc
import win32com.client as win32
import logging
from datetime import datetime
import time

# Configure the logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('AzureFunction')

def main():
    while True:
        # Execute the MS Access code
        try:
            send_sms_gateway()
        except Exception as e:
            logger.error(f'An error occurred: {str(e)}')
        
        # Wait for 10 minutes
        time.sleep(60)

def send_sms_gateway():
    # Initialize Outlook
    outlook = win32.Dispatch('Outlook.Application')

    # Connect to the MS Access database
    db_path = r'C:\Users\Jules Uel Evan Boldo\Documents\Jackfruit Maturity.accdb'
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};'
    cnxn = pyodbc.connect(conn_str)
    cursor = cnxn.cursor()

    # Execute the query
    strSQL = "SELECT Farmer, FruitCode, MaturityStatus, DaysOld FROM JackfruitMaturityQ1DummyMail WHERE DaysOld < 150 ORDER BY FruitCode;"
    cursor.execute(strSQL)

    # Loop through the records and build the email body
    body = ""
    for row in cursor:
        body += f"Fruit #{row.FruitCode} ({row.Farmer}): {row.DaysOld} days old - {row.MaturityStatus}\n"

    cursor.close()
    cnxn.close()

    # Send SMS via Email Gateway
    if body:
        # Combine all fruits in a single message
        body = f"Dear farmer,\n\nThese are the current status of the jackfruits in your farm:\n\n{body}\nPlease make sure to harvest them in time to ensure the best quality.\n\nBest regards,\nAutomated Jackfruit Maturity Monitoring System"

        # Send SMS via Email Gateway
        sms_address = "09606620507@sms.clicksend.com"
        mail_item = outlook.CreateItem(0)
        mail_item.To = sms_address
        mail_item.Subject = ""  # leave subject blank to avoid including it in the SMS message
        mail_item.Body = body
        mail_item.Send()

        # Send email to the selected recipient
        mail_item = outlook.CreateItem(0)
        mail_item.To = "julesevan2022@gmail.com"
        mail_item.Subject = f"Harvest reminder for your fruits as of {datetime.now()}"
        mail_item.Body = body
        mail_item.Send()

if __name__ == '__main__':
    main()
