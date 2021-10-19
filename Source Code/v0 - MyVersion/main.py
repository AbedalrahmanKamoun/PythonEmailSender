import xlrd, pandas, smtplib, os, os.path, time, openpyxl, re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from string import Template
import style as S

def get_recipients(recipients_filepath):
    names_list = []
    emails_list = []
    Emails_File_pandas = pandas.ExcelFile(recipients_filepath)
    Excel_File = openpyxl.load_workbook(recipients_filepath)
    Emails_Sheet_pandas = Emails_File_pandas.parse(Excel_File.sheetnames[0])
    Emails_File = xlrd.open_workbook(recipients_filepath)
    Emails_Sheet = Emails_File.sheet_by_name(Excel_File.sheetnames[0])              
    row = Emails_Sheet.nrows
    for i in range(0, row -1):
        emails_list.append(Emails_Sheet_pandas['Email'][i])
        names_list.append(Emails_Sheet_pandas['Name'][i])
    return names_list, emails_list

def get_pdfs(pdf_filepath): 
    pdf_paths_list = []
    for pdf in os.listdir(pdf_filepath):
        pdf_paths_list.append(pdf_filepath + "\\" + pdf)
    def digit(text):
        return int(text) if text.isdigit() else text
    def natural_keys(text):
        return [digit(c) for c in re.split(r'(\d+)', text)]
    pdf_paths_list.sort(key=natural_keys)
    return pdf_paths_list

def read_message_template(filename):
    with open(filename, "r", encoding="UTF-8") as message_file:
        message_file_content = message_file.read()
    return Template(message_file_content)

def smtp_connect(username, password):
    try:
        server = smtplib.SMTP(host='smtp.gmail.com', port=587)
        server.starttls()
        server.login(username, password)
        return server
    except:
        S.Notification.config(text = 'Username and Password Not Accepted!', fg = 'red')

def Finisher_msg():
    try:
        username = S.temp_sender.get()
        password = S.temp_password.get()
        server = smtp_connect(username, password)
        recipients_filepath = S.temp_receiver.get()
        names, emails = get_recipients(recipients_filepath)
        msg_path = S.temp_msg.get()
        msg_template = read_message_template(msg_path)
        msg_subject = S.temp_subject.get()
        i = 0
        while i < len(emails):
            msg = MIMEMultipart()
            msg["From"]= username
            msg["To"]= emails[i]
            msg["Subject"]= msg_subject
            msg_with_name = msg_template.substitute(student_name=names[i])
            msg.attach(MIMEText(msg_with_name))
            msg_string = msg.as_string()
            try:
                server.sendmail(username, emails[i], msg_string)
                print("Mail sent to", emails[i] + "!")
                i = i + 1
                print("*" * 60)
                print(f"{i} emails sent!")
            except Exception:
                time.sleep(5)
                server = smtp_connect(username, password)
        S.Notification.config(text = 'E-mails Successfully Sent To All Students!', fg = 'green')
        server.quit()
    except:
        S.Notification.config(text = 'Error Sending E-mails!', fg = 'red')

def Finisher_attachment():
    username = S.temp_sender.get()
    password = S.temp_password.get()
    server = smtp_connect(username, password)
    recipients_filepath = S.temp_receiver.get()
    names, emails = get_recipients(recipients_filepath)
    msg_filename = S.temp_msg.get()
    msg_template = read_message_template(msg_filename)
    msg_subject = S.temp_subject.get()
    pdf_filepath = S.temp_attachment.get()
    if os.path.isdir(pdf_filepath) == True:
        try:
            pdf_filepaths = get_pdfs(pdf_filepath)
            Attachment_name = S.temp_attach_name.get()
            i = 0
            while i < len(emails):
                msg = MIMEMultipart()
                msg_with_name = msg_template.substitute(student_name=names[i])
                msg["From"]= username
                msg["To"]= emails[i]
                msg["Subject"]= msg_subject
                msg.attach(MIMEText(msg_with_name)) 
                with open(pdf_filepaths[i], "rb") as f:
                    attach = MIMEApplication(f.read(), _subtype="pdf")
                attach.add_header("Content-Disposition", "attachment", filename=Attachment_name)
                msg.attach(attach)
                msg_string = msg.as_string()
                try:
                    server.sendmail(username, emails[i], msg_string)
                    print("Mail sent to", emails[i] + "!")
                    i = i + 1
                    print(f"{i} emails sent!")
                    print("*" * 60)
                except Exception:
                    time.sleep(5)
                    server = smtp_connect(username, password)
            S.Notification.config(text = 'E-mails Successfully Sent To All Students, Each With Its Attachment!', fg = 'green')
            server.quit()
        except:
            S.Notification.config(text = 'Error Sending E-mails!', fg = 'red')
    elif os.path.isfile(pdf_filepath) == True:
        try:
            Attachment_name = S.temp_attach_name.get()
            i = 0
            while i < len(emails):
                msg = EmailMessage()
                msg_with_name = msg_template.substitute(student_name=names[i])
                body = msg_with_name
                msg['subject'] = msg_subject
                msg['from'] = username
                msg['to'] = emails[i]
                msg.set_content(body)
                filename = S.attachments[0]
                filetype = filename.split('.')[1]
                if filetype == "jpg" or filetype == "JPG" or filetype == "png" or filetype == "PNG":
                    import imghdr
                    with open(filename, 'rb') as f:
                        file_data = f.read()
                        image_type = imghdr.what(filename)
                    msg.add_attachment(file_data, maintype='image', subtype=image_type, filename=Attachment_name)
                elif filetype == "pdf" or filetype == "PDF":
                    with open(filename, 'rb') as f:
                        file_data = f.read()
                    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=Attachment_name)
                else:
                    with open(filename, 'rb') as f:
                        file_data = f.read()
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=Attachment_name)
                try:
                    server = smtplib.SMTP('smtp.gmail.com',587)
                    server.starttls()
                    server.login(username, password)
                    server.send_message(msg)
                    print("Mail sent to", emails[i] + "!")
                    i = i + 1
                    print(f"{i} emails sent!")
                    print("*" * 60)
                except Exception:
                    time.sleep(5)
                    server = smtp_connect(username, password)
            S.Notification.config(text = 'E-mails Successfully Sent To All Students With The Selected Attachment!', fg = 'green')
            server.quit()
        except:
            S.Notification.config(text = 'Error Sending E-mails!', fg = 'red')