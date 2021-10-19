from tkinter import *
from tkinter import filedialog
import os, os.path
import main

app = Tk()
app.title("Python Email Sending App")
app.geometry("560x570")
app.config(bg = "white")

def change_yesbtn_col():
    YES_button.config(bg = "green")
    NO_button.config(bg = "red")
    ONE_button.config(state = "normal")
    MULTIPLE_button.config(state = "normal")
    Send_Attachment_button.config(state = "normal")
    Send_jsutmsg_button.config(state = "disabled")
    Attachment_name_entry.config(state = "normal")

def change_nobtn_col():
    YES_button.config(bg = "red")
    NO_button.config(bg = "green")
    ONE_button.config(state = "disabled")
    MULTIPLE_button.config(state = "disabled")
    Send_Attachment_button.config(state = "disabled")
    Send_jsutmsg_button.config(state = "normal")
    Attachment_name_entry.config(state = "disabled")

def change_onebtn_col():
    ONE_button.config(bg = "green")
    MULTIPLE_button.config(bg = "red")

def change_multiplebtn_col():
    ONE_button.config(bg = "red")
    MULTIPLE_button.config(bg = "green")

attachments = []
def create_one_attachment_entry():
    Attachment_path.place(x=15, y=270)
    Attachment_path_entry.place(x=15, y=290)
    def browse_one_attachment(): 
        filename = filedialog.askopenfilename(initialdir="c:/",title="Please Select The Attachment")
        Attachment_path_entry.delete(0, "end")
        Attachment_path_entry.insert(0, filename)
        attachments.append(filename)
    Browse_attachment_button = Button(text = "Browse File", bg = "#0563bb", fg = "white", command = browse_one_attachment)
    Browse_attachment_button.place(x=475, y=288)

def create_Multiple_attachments_entry():
    Attachment_path.place(x=15, y=270)
    Attachment_path_entry.place(x=15, y=290)
    def browse_multiple_attachment():
        filename = filedialog.askdirectory(initialdir="c:/", title="Please Select The Attachment File Directory")
        Attachment_path_entry.delete(0, "end")
        Attachment_path_entry.insert(0, filename)
    Browse_attachment_button = Button(text = "Browse File", bg = "#0563bb", fg = "white", command = browse_multiple_attachment)
    Browse_attachment_button.place(x=475, y=288)

def browse_file():
    filename = filedialog.askopenfilename(initialdir="c:/", title="Please Select The Excel File Containing The E-mails")
    Excel_file_entry.delete(0, "end")
    Excel_file_entry.insert(0, filename)

def browse_msg_text():
    filename = filedialog.askopenfilename(initialdir="c:/", title="Please Select The Message txt File")
    Message_text_entry.delete(0, "end")
    Message_text_entry.insert(0, filename)

def reset_all():
    Email_sender_entry.delete(0, "end")
    Excel_file_entry.delete(0, "end")
    Email_password_entry.delete(0, "end")
    Email_subject_entry.delete(0, "end")
    Attachment_name_entry.delete(0, "end")
    Attachment_path_entry.delete(0, "end")
    Message_text_entry.delete(0, "end")
    Notification.config(text = "")
    YES_button.config(bg = "#0563bb")
    NO_button.config(bg = "#0563bb")
    ONE_button.config(state = "disabled")
    MULTIPLE_button.config(state = "disabled")
    ONE_button.config(bg = "#0563bb")
    MULTIPLE_button.config(bg = "#0563bb")
    Send_Attachment_button.config(state = "disabled")
    Send_jsutmsg_button.config(state = "disabled")

# HEADER
heading = Label(text = 'تطبيق جمعيّة "تأسيسيّة" لارسال البرائد الالكترونيّة', bg = "#0563bb", fg = "white", font = "20", width = "500", height = "2")
heading.pack()
subheading = Label(text = "Please fill the form below to send the E-mail", bg = "white", fg = "#0563bb", font = "10")
subheading.pack()

# LABELS
Email_sender = Label(text = "Enter User E-mail Address:", bg = "white")
Email_sender.place(x=15, y=90)
Email_password = Label(text = "Enter User Password:", bg = "white")
Email_password.place(x=300, y=90)
Excel_file = Label(text = "Insert The Recipents' E-mails Excel File:", bg = "white")
Excel_file.place(x=15, y=140)
Attachment_question = Label(text = "Are You Sending An Attachment?", bg = "white")
Attachment_question.place(x=15, y=190)
Attachment_quantity_question = Label(text = "If YES Select the Proper Option:", bg = "white")
Attachment_quantity_question.place(x=295, y=190)
Attachment_path = Label(text = "Insert The Attachment File Or Directory:", bg = "white") # place in function
Message_text = Label(text = "Insert The Message Template Path:", bg = "white")
Message_text.place(x=15, y=320)
Email_subject = Label(text = "Enter E-mail Subject:", bg = "white")
Email_subject.place(x=15, y=370)
Attachment_name = Label(text = "Enter The Attachment Name:", bg = "white")
Attachment_name.place(x=300, y=370)

# ENTRIES
temp_sender = StringVar()
temp_password = StringVar()
temp_receiver = StringVar()
temp_msg = StringVar()
temp_subject = StringVar()
temp_body = StringVar()
temp_attachment = StringVar()
temp_attach_name = StringVar()
Email_sender_entry = Entry(textvariable = temp_sender, width = "40")
Email_password_entry = Entry(textvariable = temp_password, width = "40", show = "•")
Excel_file_entry = Entry(textvariable = temp_receiver, width = "75", fg = "green")
Attachment_path_entry = Entry(textvariable = temp_attachment, width = "75", fg = "green") # place in function
Message_text_entry = Entry(textvariable = temp_msg, width = "75", fg = "green")
Email_subject_entry = Entry(textvariable = temp_subject, width = "40")
Attachment_name_entry = Entry(textvariable = temp_attach_name, width = "40", state = "disabled")
Email_sender_entry.place(x=15, y=110)
Email_password_entry.place(x=300, y=110)
Excel_file_entry.place(x=15, y=160)
Message_text_entry.place(x=15, y=340)
Email_subject_entry.place(x=15, y=390)
Attachment_name_entry.place(x=300, y=390)

# BUTTONS
YES_button = Button(text = "YES", bg = "#0563bb", fg = "white", width = "5", command = change_yesbtn_col)
YES_button.place(x= 90, y= 210)
NO_button = Button(text = "NO", bg = "#0563bb", fg = "white", width = "5", command = lambda: [change_nobtn_col()])
NO_button.place(x= 90, y= 240)
ONE_button = Button(text = "Single Attachment To All Students", fg = "white", bg = "#0563bb", state = "disabled", command = lambda: [create_one_attachment_entry(), change_onebtn_col()])
ONE_button.place(x= 300, y= 210)
MULTIPLE_button = Button(text = "Each Student Will Get His Own Attachment", fg = "white", bg = "#0563bb", state = "disabled", command = lambda: [create_Multiple_attachments_entry(), change_multiplebtn_col()])
MULTIPLE_button.place(x=300, y= 240)
Browse_Excel_button = Button(text = "Browse File", bg = "#0563bb", fg = "white", command = browse_file)
Browse_Excel_button.place(x=475, y=158)
Browse_msg_button = Button(text = "Browse File", bg = "#0563bb", fg = "white", command = browse_msg_text)
Browse_msg_button.place(x=475, y=338)
Send_jsutmsg_button = Button(text = "Send E-mails With No Attachments", bg = "#0563bb", fg = "white", width = "35", height = "2", state = "disabled", command = main.finisher_msg)
Send_jsutmsg_button.place(x=20, y=430)
Send_Attachment_button = Button(text = "Send E-mails + Attachments", bg = "#0563bb", fg = "white", width = "35", height = "2", state = "disabled", command = main.finisher_attachment)
Send_Attachment_button.place(x=20, y=470)
Reset_button = Button(text = "Reset The Form", bg = "#0563bb", fg = "white", width = "35", height = "2", command = reset_all)
Reset_button.place(x=285, y=430)

# NOTIFICATION
Notification = Label(text="", font = "20", bg = "white", fg = "green")
Notification.place(x=30, y=520)

# FOOTER
Footer = Label(text = "This Application is Developed by Ibrahim Maassarani, Abed-AlRahman Kamoun & Issa Kassas", bg = "white")
Footer.place(x=20, y=550)

app.mainloop()