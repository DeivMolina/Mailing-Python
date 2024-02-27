import pandas as pd
import smtplib

# Set credentials
name_account = "Milton Cruz"
email_account = "miltoncn56720@gmail.com"
password_account = "srukvrpjvpwnvkkg" # password created in step 1.3.

# 'smtp.gmail.com' and 465 port refer to Gmail as provider
# Change these arguments if you are using another one
# For example, Outlook arguments are 'smtp-mail.outlook.com' and 587 port
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(email_account, password_account)

# Read the file that contains at least names & email addresses
# Subjects & messages can be personalized, but we use them as input
email_df = pd.read_excel("Data/Emails.xlsx")

# Get all names, email addresses, subjects & messages
all_names = email_df['Name']
all_emails = email_df['Email']
all_subjects = email_df['Subject']
all_messages = email_df['Message']

for i in range(len(email_df)):

    name = all_names[i]
    email = all_emails[i]
    
    # Personalized subject
    subject = all_subjects[i] + ', ' + all_names[i] + '!'
    
    # Personalized message
    message = ('Hey, ' + all_names[i] + '!\n\n' +
              all_messages[i] + '\n\n'
              'Te deseamos lo mejor,\n' +
              name_account)

    # Generate email to be sent
    sent_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(name_account, email_account, name, email, subject, message))
    
    # Send emails, the reason will be displayed in case you get an error
    try:
        server.sendmail(email_account, [email], sent_email) # [email] is a list that can contain multiple emails
    except Exception:
        print('Could not send email to {}. Error: {}\n'.format(email, str(Exception)))

# Close smtp server
server.close()