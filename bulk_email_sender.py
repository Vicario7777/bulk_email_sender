import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Read SMTP server credentials from a separate config file or environment variables
SMTP_SERVER = 'smtp.example.com'
SMTP_PORT = 587
SMTP_USERNAME = 'your_username'
SMTP_PASSWORD = 'your_password'

# Load data from Exel Sheet
try:
    excel_data = pd.read_excel('Data for customer bulk emails.xlsx', sheet_name='Customers')
except FileNotFoundError:
    print("Excel file not found.")
    exit()

# Set up the SMTP connection
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(SMTP_USERNAME, SMTP_PASSWORD)
except smtplib.SMTPAuthenticationError:
    print("SMTP authentication failed.")
except smtplib.SMTPException as e:
    print("An error occurred while connecting to the SMTP server", e)
    exit()

# Iterate through the data and send emails
for index, row in excel_data.iterrows():
    try:
        # Create a MIME object for the email
        msg = MIMEMultipart()
        msg['From'] = SMTP_USERNAME
        msg['To'] = row['Email']
        msg['Subject'] = 'Your SUbject Here'

        # Customise the email body
        message = f"Dear {row['Name']},\n\n"
        message += "Your customised message here.\n\n"
        message += "Warmest Regards,\nYour Name"
        msg.attach(MIMEText(message, 'plain'))

        # Send the email
        server.sendmail(SMTP_USERNAME, row['Email'], msg.as_string())
        print(f"Email sent to {row['Email']}")
    except KeyError as e:
        print("KeyError occurred. Missing column in Excel data:", e)
    except Exception as e:
        print("An error occurred while sending email to", row['Email'], ":", e)

# Close the SMTP connection
server.quit()
              