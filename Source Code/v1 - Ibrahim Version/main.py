import xlrd, pandas, smtplib, getpass, os, os.path, time, openpyxl, re
from os import path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.message import EmailMessage
from string import Template
import style as S
import imghdr

image_formats_list = ["jpeg", "jpg", "png"] # initialized as global variable for possible future modifications

def get_recipients_info(recipients_filepath):
    """ Fetch the names and emails of recipients from an Excel Sheet

    Parameters
    ----------
    recipients_filepath: {string}, filepath of the Excel Sheet
    
    Returns
    -------
    names_list: {list}, full names of the recipients from the Excel Sheet
    emails_list: {list}, emails of the recipients from the Excel Sheet

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
        emails_list.append(emails_sheet_pandas["Email"][i])
        names_list.append(emails_sheet_pandas["Name"][i])
    return names_list, emails_list

def get_attachments(attachment_filepath):
    """ Fetch the full path of the file(s) to be attached

    Parameters
    ----------
    attachment_filepath: {string}, filepath of a single attachment or the directory that contains one or more attachments
    
    Returns
    -------
    pdf_paths_list: {list}, full path to every attachment (could be one attachment, and hence, one path)
    
    """
    def digit(text):
        return int(text) if text.isdigit() else text
    def natural_keys(text):
        return [digit(c) for c in re.split(r"(\d+)", text)]
    pdf_paths_list = []
    for pdf in os.listdir(attachment_filepath):
        pdf_paths_list.append(attachment_filepath + "\\" + pdf)
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
        server = smtplib.SMTP(host="smtp.gmail.com", port=587)
        server.starttls()
        server.login(username, password)
        return server
    except:
        S.Notification.config(text = "Username and Password Not Accepted!", fg = "red")

def smtp_reconnect(username, password):
    """ Reopens a new connection to the Gmail SMTP server if the old connection brakes

    Parameters
    ----------
    username: {string}, same in smtp_connect
    password: {string}, same in smtp_connect
    
    Returns
    -------
    server: {object}, same in smtp_connect
    
    """
    time.sleep(5)
    server = smtp_connect(username, password)
    return server

def get_parameters():
    """ Get the necessary parameters for connecting to the SMTP server, and sending emails with/without attachment(s)

    Parameters
    ----------
    No parameters are taken
    
    Returns
    -------
    username: {string}, user input from the GUI (parameter in smtp_connect)
    password: {string}, user input from the GUI (parameter in smtp_connect)
    server: {object} (returned in smtp_connect)
    names, emails: {list} (returned in get_recipients_info)
    msg_template: {object} (returned in read_message_template)
    msg_subject: {string}, the subject of the email, user input from the GUI
    
    """
    username = S.temp_sender.get() 
    password = S.temp_password.get()
    server = smtp_connect(username, password)
    recipients_filepath = S.temp_receiver.get()
    names, emails = get_recipients_info(recipients_filepath)
    msg_path = S.temp_msg.get()
    msg_template = read_message_template(msg_path)
    msg_subject = S.temp_subject.get()
    return username, password, server, names, emails, msg_template, msg_subject

def finisher_msg():
    """ Sends emails without attachments

    Parameters
    ----------
    No parameters are taken
    
    Returns
    -------
    Nothing is returned
    
    """
    username, password, server, names, emails, msg_template, msg_subject = get_parameters()
    i = 0
    try:
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
                print("Email Sent to", emails[i] + "!")
                i = i + 1
                print(f"{i} Emails Sent!")
                print("*" * 60)
            except Exception:
                server = smtp_reconnect(username, password)
        S.Notification.config(text = "Emails Successfully Sent to All Students!", fg = "green")
        server.quit()
    except:
        S.Notification.config(text = "Error Sending Emails!", fg = "red")

def finisher_attachment():
    """ Sends emails with attachment(s)

    Parameters
    ----------
    No parameters are taken
    
    Returns
    -------
    Nothing is returned
    
    """
    username, password, server, names, emails, msg_template, msg_subject = get_parameters()
    attachment_filepath = S.temp_attachment.get()
    if os.path.isdir(attachment_filepath) == True:
        try:
            attachment_filepaths = get_attachments(attachment_filepath)
            attachment_name = S.temp_attach_name.get()
            i = 0
            while i < len(emails):
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
                    i = i + 1
                    print(f"{i} Emails Sent!")
                    print("*" * 60)
                except Exception:
                    server = smtp_reconnect(username, password)
            S.Notification.config(text = "Emails Successfully Sent to all Students, Each with its Attachment!", fg = "green")
            server.quit()
        except:
            S.Notification.config(text = "Error Sending Emails!", fg = "red")
    elif os.path.isfile(attachment_filepath) == True:
        try:
            attachment_name = S.temp_attach_name.get()
            i = 0
            while i < len(emails):
                msg = EmailMessage()
                msg_with_name = msg_template.substitute(student_name=names[i])
                msg["subject"] = msg_subject
                msg["from"] = username
                msg["to"] = emails[i]
                msg.set_content(msg_with_name)
                filename = S.attachments[0]
                file_type = filename.split(".")[1].lower()
                with open(filename, "rb") as f:
                    file_data = f.read()
                    main_type = "application"
                    if file_type in image_formats_list:
                        file_type = imghdr.what(filename)
                        main_type = "image"
                    elif file_type != "pdf":
                        file_type = "octet-stream"
                msg.add_attachment(file_data, maintype=main_type, subtype=file_type, filename=attachment_name)
                try:
                    server.send_message(msg)
                    print("Email Sent to", emails[i] + "!")
                    i = i + 1
                    print(f"{i} Emails Sent!")
                    print("*" * 60)
                except Exception:
                    server = smtp_reconnect(username, password)
            S.Notification.config(text = "Emails Successfully Sent to all Students with the Selected Attachment!", fg = "green")
            server.quit()
        except:
            S.Notification.config(text = "Error Sending Emails!", fg = "red")