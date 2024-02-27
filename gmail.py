import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

# Set credentials (consider using a more secure method for storing credentials)
name_account = "Owner"
email_account = "owner@dm-series.com"
password_account = "Dgm2021*#/-"

# 'smtp.gmail.com' and 465 port refer to Gmail as a provider
# Change these arguments if you are using another one
# For example, Outlook arguments are 'smtp-mail.outlook.com' and 587 ports
server = smtplib.SMTP_SSL('smtp.hostinger.com', 465)
server.ehlo()
server.login(email_account, password_account)

# Read the file that contains at least names & email addresses
# Subjects & messages can be personalized, but we use them as input
email_df = pd.read_excel("Data/Emails.xlsx")

def send_email(name, email, subject, html_message, cc_email="owner@dm-series.com")  :
    try:
        mime_message = MIMEMultipart()
        mime_message.attach(MIMEText(html_message, 'html'))
        mime_message['From'] = "{} <{}>".format(name_account, email_account)
        mime_message['To'] = "{} <{}>".format(name, email)
        mime_message['Subject'] = subject

        # Agrega la dirección de correo para copia (CC) si se proporciona
        if cc_email:
            mime_message['Cc'] = cc_email

        recipients = [email]
        if cc_email:
            recipients.append(cc_email)

        server.sendmail(email_account, recipients, mime_message.as_string())
        print(f'Successfully sent email to {email} with CC to {cc_email}')

    except smtplib.SMTPException as e:
        print(f'Error sending email to {email} with CC to {cc_email}: {e}')


def generate_html_message(name, message, sender_name, template_path='template.html'):
    # Verifica si el archivo de plantilla existe
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'Template file not found: {template_path}')

    # Lee el contenido del archivo de plantilla
    with open(template_path, 'r', encoding='utf-8') as file:
        template_content = file.read()

    # Sustituye las variables en la plantilla
    html_message = template_content.format(name=name, message=message, sender_name=sender_name)
    return html_message


# Get all names, email addresses, subjects & messages
all_names = email_df['Name']
all_emails = email_df['Email']
all_subjects = email_df['Subject']
all_messages = email_df['Message']

try:
    for i in range(len(email_df)):
        name = all_names[i]
        email = all_emails[i]
        subject = f'{all_subjects[i]}, {all_names[i]}!'
        html_message = generate_html_message(name, all_messages[i], name_account)

        send_email(name, email, subject, html_message)

except Exception as e:
    print(f'Error sending emails: {e}')

finally:
    server.quit()
