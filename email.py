import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Path to the Excel file
excel_file = "mailsheet.xlsx"  # Adjust this path if the file is in a different location
cv_file_path = "kifayat_aliResume.pdf"
# SMTP configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "kifayat.siliconplex@gmail.com"
app_password = "dhhv vlzl ozbt iegb"  # Use an App Password if 2FA is enabled

# Load Excel data
df = pd.read_excel(excel_file)


# Function to send email
def send_email(receiver_email, receiver_name):
    try:
        # Set up the server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(sender_email, app_password)

        # Compose email
        email_message = MIMEMultipart()
        email_message["From"] = sender_email
        email_message["To"] = receiver_email
        email_message["Subject"] = "Request for acceptance letter for CSC Scholarship "
        
        # Customize the message with the recipient's name
        message_body = f"""\
Dear Professor {receiver_name},

I hope this email finds you well. My name is Kifayat Ali, and I am from Pakistan. I completed my Bachelor's degree in Information Technology in 2020 and have since been working as a teacher and backend developer, specializing in Python and PHP.

I am writing to express my strong interest in pursuing a Master's degree at your esteemed institution under your supervision through the Chinese Government Scholarship (CSC). I am particularly drawn to your research areas, which perfectly align with my academic interests and career goals. I am eager to contribute to your work in these critical fields while enhancing my knowledge and skills.

I would be honored if you could consider my application for supervision and provide me with an acceptance letter for the CSC scholarship. Attached to this email is my CV for your review.

Thank you for your time and consideration. I look forward to the possibility of working under your guidance.

Best regards,

Kifayat Ali
+923469238735
"""

        # Attach the CV

        email_message.attach(MIMEText(message_body, "plain"))
        with open(cv_file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={cv_file_path.split('/')[-1]}")
            email_message.attach(part)


        # Send the email
        server.sendmail(sender_email, receiver_email, email_message.as_string())
        print(f"Email sent successfully to {receiver_email}")
    
    except smtplib.SMTPAuthenticationError as e:
        print("Error connecting to the email server:", e)
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {e}")
    finally:
        server.quit()

# Loop through each row in the Excel sheet and send an email
for index, row in df.iterrows():
    receiver_name = row["name"]
    receiver_email = row["email"]
    send_email(receiver_email, receiver_name)
