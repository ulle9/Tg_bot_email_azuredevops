from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import ssl
from imapclient import IMAPClient
from datetime import datetime
import configparser


config = configparser.ConfigParser()
config.read_file(open(r'e_coll.cfg', encoding='utf-8'))
smtp_host = config.get('email_worker', 'smtp_host')
smtp_port = int(config.get('email_worker', 'smtp_port'))
imap_host = config.get('email_worker', 'imap_host')
imap_port = int(config.get('email_worker', 'imap_port'))
project_dict = dict()
for i in config.get('email_worker', 'project_dict').split(';'):
    project_dict[i.split(':')[0]] = i.split(':')[1].split(',')


# Email body parser
def parse_email_body(email_body, length=300):
    body = ""
    if email_body.is_multipart():
        for part in email_body.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))
            if ctype == 'text/plain' and 'attachment' not in cdispo:
                body = part.get_payload(decode=True)  # decode
                if part.get_content_charset():
                    body = body.decode(part.get_content_charset(), 'ignore')[:length]
                    break
                else:
                    body = body.decode('cp1251', 'ignore')[:length]

    else:
        body = email_body.get_payload(decode=True)  # decode
        if email_body.get_content_charset():
            body = body.decode(email_body.get_content_charset(), 'ignore')[:length]
        else:
            body = body.decode('unicode_escape', 'ignore')[:length]
    return body


# Send email function
def send_email(project, msg_date, reply_msg, reg_item_num):
    msg = MIMEMultipart()
    msg['From'] = project_dict[project][2]
    msg['To'] = project_dict[project][4]
    msg['Subject'] = "{}. Text {} text {}".format(project_dict[project][0], reg_item_num, msg_date)
    message = 'Text {} text {}text{}\n\n\n'.format(reg_item_num, msg_date, project_dict[project][1]) + \
              'Text' + '-' * 200 + '\n' + reply_msg
    msg.attach(MIMEText(message, 'plain'))

    # Create server
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(smtp_host, port=smtp_port, context=context)
    # Login
    server.login(project_dict[project][2], project_dict[project][3])
    # Send the message via the server
    rcpt = [msg['To'], project_dict[project][5]]
    server.sendmail(msg['From'], rcpt, msg.as_string())
    server.quit()

    # Put email to 'Sent' folder
    msg['Bcc'] = project_dict[project][5]
    msg['Date'] = str(datetime.now().strftime("%Y %b %d %H:%M:%S"))
    text = msg.as_string()
    client = IMAPClient(imap_host, port=imap_port, use_uid=True, ssl=True)
    client.login(project_dict[project][2], project_dict[project][3])
    client.append('Sent', text, '\\Seen', datetime.now())
    client.logout()

