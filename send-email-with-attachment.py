from datetime import datetime, timedelta
import smtplib
import openpyxl
from email.message import EmailMessage
import time



#the below line is used to create an attachment on the email
attachment = './data_csv/orders_data-2017-2021.csv'


# Create an email Message
def create_message(to_add, from_add, subject, body, cc=None):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = from_add
    msg['CC'] = cc
    msg['To'] = to_add
    #msg['X-Priority'] = '2'
    msg.add_alternative(body, subtype='html')
    return msg

# Send email and add attachment
def send_email(to_add, from_add, msg, attachment):
    with open(attachment, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename='orders-data-2017-2021.csv')
    with smtplib.SMTP('', 25) as smtp: #add here the Simple Mail Transfer Protocol (SMTP)
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.send_message(msg)
        smtp.close


start_time = time.time()

username = 'John'
cc_email = 'myemail@myemail.com' #add here any email address you would like to add in the cc list
user_email = 'gyouremail@youremail.com' #add here the email address of the intended receiver

# Email body: Use this link to convert your text to HTML https://www.textfixer.com/html/convert-text-html.php
BODY = f"""
        <!DOCTYPE html>
            <html>
                <body>
                    <p><i>This email has been sent automatically via a python script.</i></p>
                    <p>Hello {username}! - I hope this finds you well :) </p>
                    <p>Please find attached the a table summarizing the orders for the last 5 years</p>
                    <p>Thank you,</p>
                </body>
            </html>
        """

# Complie the email message
                    #TO ADD ; FROM ADD ; SUBJ
msg = create_message(user_email,
                     'myemail@myemail.com', #add here email of the sender
                     "Last 5 years orders data", #subject of the email
                     BODY, #pass here the body of the email
                     cc_email) #add here the list created above that contains all the addresses you want to include in the cc list
# Send Email
send_email(user_email, 'myemail@myemail.com', msg, attachment)

# Get the run time of the code
duration = f"{time.time() - start_time} seconds"
print(duration)