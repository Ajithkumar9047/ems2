import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib
from datetime import datetime, timedelta
import os
from Config import *
from dotenv import load_dotenv
load_dotenv()
current_date1 = datetime.now().strftime('%d/%m/%Y')
yesterday_date1 = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')

def get_cc_from_email(msg):
    cc_header = msg.get("Cc", "")
    return [cc.strip() for cc in cc_header.split(",") if cc]

def email_send(attachment_path):
    def reply_all_with_attachment_and_cc(email_subject, reply_text, attachment_path, imap_server, imap_username, imap_password, smtp_server, smtp_port, smtp_username, smtp_password):
        try:
            with imaplib.IMAP4_SSL(imap_server) as mail:
                mail.login(imap_username, imap_password)
                mail.select("inbox")
                result, data = mail.search(None, f'SUBJECT "{email_subject}"')

                if result == "OK":
                    # Get the list of email IDs
                    email_ids = data[0].split()
                    if email_ids:
                        # Loop through the email IDs (considering the most recent one)
                        for email_id in reversed(email_ids):
                            # Fetch the email with the specified ID
                            result, message_data = mail.fetch(email_id, "(RFC822)")

                            if result == "OK":
                                # Parse the email content
                                msg = email.message_from_bytes(message_data[0][1])

                                # Extract relevant information
                                from_email ='rajesh.yerremshetty@unicorntechmedia.com'
                                to_email = 'ajithkumar@newgendigital.com'
                                #cc_email_list = ccMail
                                cc_email_str = os.getenv("CC_EMAIL")
                                cc_email_list = cc_email_str.strip("[]").replace("'", "").split(", ")
                                subject = f"{msg['Subject']}"

                                # Create the reply message
                                reply_message = MIMEMultipart()
                                reply_message["From"] = to_email
                                reply_message["To"] = from_email
                                reply_message["Cc"] = ", ".join(cc_email_list)
                                reply_message["Subject"] = subject
                                reply_message.attach(MIMEText(reply_text, "html"))
                                original_attachment_name = os.path.basename(attachment_path)
                                new_excel_file_name = f"Ems_Repush_{yesterday_date1.replace('/', '_')}AM_to_{current_date1.replace('/', '_')}.{original_attachment_name.split('.')[-1]}"

                                excel_attachment = open(attachment_path, 'rb')
                                excel_part = MIMEApplication(excel_attachment.read(), Name=new_excel_file_name)
                                excel_attachment.close()
                                excel_part['Content-Disposition'] = f'attachment; filename="{new_excel_file_name}"'
                                reply_message.attach(excel_part)
                                with smtplib.SMTP(smtp_server, smtp_port) as server:
                                    server.starttls()
                                    server.login(smtp_username, smtp_password)
                                    recipients = [to_email] + cc_email_list
                                    server.sendmail(to_email, recipients, reply_message.as_string())

                                print(f"Reply sent to {from_email} with CC to {', '.join(cc_email_list)} for the email with subject '{subject}'.")
                                logging.info(f"Reply sent to {from_email} with CC to {', '.join(cc_email_list)} for the email with subject '{subject}'.")
                                break  # Only reply to the most recent email for simplicity
                            else:
                                print(f"Failed to fetch email with ID {email_id}.")
                                logging.info(f"Failed to fetch email with ID {email_id}.")
                    else:
                        print(f"No emails found with subject '{email_subject}'.")
                        logging.info(f"No emails found with subject '{email_subject}'.")
                else:
                    print(f"Error searching for emails with subject '{email_subject}': {data}")
                    logging.info(f"Error searching for emails with subject '{email_subject}': {data}")

        except Exception as e:
            print(f"An error occurred: {str(e)}")
            logging.error(f"An error occurred: {str(e)}")

    email_subject = "Re: Updated EMS response for null EMS"
    reply_text = f'''\
        <html>
        <head>
        <style>
            body {{
                font-family: Verdana, sans-serif;
                font-size: 14px;
                color: black;
            }}
            p {{
                margin: 10px 0;
            }}
        </style>
        </head>
        <body>
        <p>Hi all,</p>

        <p>Please find the updated EMS responses for the leads where you received null as EMS response.</p>
        <p>Note: These EMS responses are for the leads that we received between {yesterday_date1} 12:00AM and {current_date1} 6:00PM</p>

        <b style="color: blue;">Thanks & Regards,<br></b>
        <p>Ajithkumar Sekar | Technical support executive<br>
        Newgen Digital Works Pvt. Ltd.<br>
        M: (+91) 8072467327<br>
        W: <a href="http://www.newgen.co">www.newgen.co</a><br>
        A: Chennai - 600041.</p>
        </body>
        </html>
        '''

    imap_server = "imap.gmail.com"
    imap_username = "ajithkumar@newgendigital.com"
    imap_password = os.getenv("DB_PASSWORD")

    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_username = "ajithkumar@newgendigital.com"
    smtp_password = os.getenv("DB_PASSWORD")
    
    reply_all_with_attachment_and_cc(email_subject, reply_text, attachment_path, imap_server, imap_username, imap_password, smtp_server, smtp_port, smtp_username, smtp_password)
