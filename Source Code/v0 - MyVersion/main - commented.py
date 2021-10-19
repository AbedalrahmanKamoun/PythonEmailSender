import xlrd, pandas, smtplib, getpass, os, os.path, time, openpyxl, re
from os import path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from string import Template
import style as S

def get_recipients_info(recipients_filepath):
    """ Fetch the names and emails of recipients from an Excel Sheet

    Parameters
    ----------
    recipients_filepath: {string}, filepath of the Excel Sheet
    
    Returns
    -------
    names_list: {list}, full names of the recipients
    emails_list: {list}, emails of the recipients

    """
    names_list = []
    emails_list = []
    emails_file_pandas = pandas.ExcelFile(recipients_filepath)
    excel_file = openpyxl.load_workbook(recipients_filepath)
    emails_sheet_pandas = emails_file_pandas.parse(excel_file.sheetnames[0])
    emails_file = xlrd.open_workbook(recipients_filepath)
    emails_sheet = emails_file.sheet_by_name(excel_file.sheetnames[0])              
    row = emails_sheet.nrows
    for i in range(0, row -1):
        emails_list.append(emails_sheet_pandas['Email'][i])
        names_list.append(emails_sheet_pandas['Name'][i])
    return names_list, emails_list

def get_attachments(attachment_filepath):
    """ Fetch the full path of the file(s) to be attached

    Parameters
    ----------
    attachment_filepath: {string}, filepath of a single attachment or the directory that contains one or more attachments
    
    Returns
    -------
    pdf_paths_list: {list}, full path to the attachment(s)
    
    """
    pdf_paths_list = []
    for pdf in os.listdir(attachment_filepath):
        pdf_paths_list.append(attachment_filepath + "\\" + pdf)
    def digit(text): # why defined inside the function?
        return int(text) if text.isdigit() else text
    def natural_keys(text): # same, why defined inside the function?
        return [digit(c) for c in re.split(r'(\d+)', text)]
    pdf_paths_list.sort(key=natural_keys)
    return pdf_paths_list

def read_message_template(message_filepath):
    """ Reads the content of a .txt file to form the message template to bet sent

    Parameters
    ----------
    message_filepath: {string}, filepath of the .txt file
    
    Returns
    -------
    Template: {object}, represents the content of the message
    
    """
    with open(message_filepath, "r", encoding="UTF-8") as message_file:
        message_file_content = message_file.read()
    return Template(message_file_content)

def smtp_connect(username, password):
    """ Opens a connection to the Gmail SMTP server

    Parameters
    ----------
    username: {string}, the username, i.e., the email, of the account from which the emails are to be sent
    password: {string}, the password of the account
    
    Returns
    -------
    server: {object}, the SMTP server with an open connection if the credentials are legit
    
    """
    try:
        server = smtplib.SMTP(host='smtp.gmail.com', port=587)
        server.starttls()
        server.login(username, password)
        return server
    except:
        S.Notification.config(text = 'Username and Password Not Accepted!', fg = 'red')

def finisher_msg():
    """ Sends email without attachments

    Parameters
    ----------
    No parameters are taken
    
    Returns
    -------
    Nothing is returned
    
    """
    try:
        # why repeated multiple times in every send option?
        username = S.temp_sender.get() 
        password = S.temp_password.get()
        server = smtp_connect(username, password)
        recipients_filepath = S.temp_receiver.get()
        names, emails = get_recipients_info(recipients_filepath)
        msg_path = S.temp_msg.get()
        msg_template = read_message_template(msg_path)
        msg_subject = S.temp_subject.get()
        # the previous lines are independent from the with/without attachment option
        for i in range(len(emails)): # use while instead
            msg = MIMEMultipart()
            msg["From"]= username
            msg["To"]= emails[i]
            msg["Subject"]= msg_subject
            msg_with_name = msg_template.substitute(student_name=names[i])
            msg.attach(MIMEText(msg_with_name))
            msg_string = msg.as_string()
            try:
                server.sendmail(username, emails[i], msg_string)
                print("Email Sent to", emails[i] + "!")
                print(f"{i+1} Emails Sent!")
                print("*" * 60)
            except Exception:
                time.sleep(5)
                server = smtp_connect(username, password)
        S.Notification.config(text = 'Emails Successfully Sent to All Students!', fg = 'green')
        server.quit()
    except:
        S.Notification.config(text = 'Error Sending Emails!', fg = 'red')

def finisher_attachment():
    """ Sends email with attachment(s)

    Parameters
    ----------
    No parameters are taken
    
    Returns
    -------
    Nothing is returned
    
    """
    username = S.temp_sender.get()
    password = S.temp_password.get()
    server = smtp_connect(username, password)
    recipients_filepath = S.temp_receiver.get()
    names, emails = get_recipients_info(recipients_filepath)
    msg_filename = S.temp_msg.get()
    msg_template = read_message_template(msg_filename)
    msg_subject = S.temp_subject.get()
    attachment_filepath = S.temp_attachment.get()
    if os.path.isdir(attachment_filepath) == True:
        try:
            attachment_filepaths = get_attachments(attachment_filepath)
            attachment_name = S.temp_attach_name.get()
            for i in range(len(emails)): # use while instead
                msg = MIMEMultipart()
                msg_with_name = msg_template.substitute(student_name=names[i])
                msg["From"]= username
                msg["To"]= emails[i]
                msg["Subject"]= msg_subject
                msg.attach(MIMEText(msg_with_name)) 
                with open(attachment_filepaths[i], "rb") as f:
                    attach = MIMEApplication(f.read(), _subtype="pdf")
                attach.add_header("Content-Disposition", "attachment", filename=attachment_name)
                msg.attach(attach)
                msg_string = msg.as_string()
                try:
                    server.sendmail(username, emails[i], msg_string)
                    print("Email Sent to", emails[i] + "!")
                    print(f"{i+1} Emails Sent!")
                    print("*" * 60)
                except Exception:
                    time.sleep(5)
                    server = smtp_connect(username, password)
            S.Notification.config(text = 'Emails Successfully Sent to all Students, Each with its Attachment!', fg = 'green')
            server.quit()
        except:
            S.Notification.config(text = 'Error Sending Emails!', fg = 'red')
    elif os.path.isfile(attachment_filepath) == True:
        try:
            attachment_name = S.temp_attach_name.get()
            for i in range(len(emails)): # use while instead
                msg = EmailMessage()
                msg_with_name = msg_template.substitute(student_name=names[i])
                body = msg_with_name # why using body?
                msg['subject'] = msg_subject
                msg['from'] = username
                msg['to'] = emails[i]
                msg.set_content(body) # why using body?
                filename = S.attachments[0]
                filetype = filename.split('.')[1] # why to reassign in another line?
                # filetype = filetype[1]

                # isn't there an issue with the logic of the conditioning? why not only test on filetype directly and replace subtype in add_attachement()
                if filetype == "jpg" or filetype == "JPG" or filetype == "png" or filetype == "PNG": # why not use if in list: if filetype in [a, b, ...]
                    import imghdr # why import library here?
                    with open(filename, 'rb') as f:
                        file_data = f.read()
                        image_type = imghdr.what(filename)
                    msg.add_attachment(file_data, maintype='image', subtype=image_type, filename=attachment_name)
                elif filetype == "pdf" or filetype == "PDF": # why not use if in list
                    with open(filename, 'rb') as f:
                        file_data = f.read()
                    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=attachment_name)
                else:
                    with open(filename, 'rb') as f:
                        file_data = f.read()
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=attachment_name)
                try:
                    server = smtplib.SMTP('smtp.gmail.com',587)
                    server.starttls()
                    server.login(username, password)
                    server.send_message(msg)
                    print("Mail Sent to", emails[i] + "!")
                    print(f"{i+1} Emails Sent!")
                    print("*" * 60)
                except Exception:
                    time.sleep(5)
                    server = smtp_connect(username, password)
            S.Notification.config(text = 'Emails Successfully Sent to all Students with the Selected Attachment!', fg = 'green')
            server.quit()
        except:
            S.Notification.config(text = 'Error Sending Emails!', fg = 'red')