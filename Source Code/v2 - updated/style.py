from tkinter import *
from tkinter import filedialog
from tkinter import messagebox as mb
import os, os.path
import signal
import main

app = Tk()
app.title("Python Email Sending App")
app.geometry("560x570")
app.config(bg="white")

def interrupt_app():
    os.kill(os.getpid(), signal.CTRL_C_EVENT) # send CTRL + C to app (only in windows)

def handle_entry_focus_red(event):
    if event.widget["bg"] == "red":
        event.widget["bg"] = "white"

def handle_button_focus_red(event):
    if event.widget["textvariable"] == "":
        event.widget["bg"] = "white"

def check_entries(entries_list):
    c = 0
    for entry in entries_list:
        if entry["state"] == "normal" and not entry.get():
            app.focus()
            entry.config(bg="red")
            entry.bind()
            c = 1
    if c == 1:
        display_empty_fields()

def display_successfull_send():
    mb.showinfo(title="Congrats!", message="All Emails have been Sent Successfully...")

def display_connection_problems():
    mb.showwarning(title="Potential Connection Problems!", message="Kindly Check your Internet Connection...")

def display_wrong_credentials():
    mb.showerror(title="Invalid Credentials!", message="Wrong Username and or Password...")

def display_path_check():
    mb.showerror(title="Invalid Path!", message="Kindly Check your Choice and the Provided Path...")

def display_unexpected_error():
    mb.showerror(title="Unexpected Error!", message="Something Went Wrong...")

def display_empty_fields():
    mb.showerror(title="Required Field!", message="Mandatory Data is Missing...")

def change_yesbtn_col():
    yes_button.config(bg="light blue")
    no_button.config(bg="red")
    one_button.config(state="normal", bg="#0563bb", selectcolor="#0563bb")
    multiple_button.config(state="normal", bg="#0563bb", selectcolor="#0563bb")
    browse_attachment_button.config(state="disabled")
    attachment_path_entry.config(state="normal")
    attachment_name_entry.config(state="normal")
    send_attachment_button.config(state="normal")
    send_jsutmsg_button.config(state="disabled")

def change_nobtn_col():
    yes_button.config(bg="red")
    no_button.config(bg="light blue")
    attachment_choice.set(0)
    one_button.config(state="disabled", bg="#0563bb", selectcolor="#0563bb")
    multiple_button.config(state="disabled", bg="#0563bb", selectcolor="#0563bb")
    browse_attachment_button.config(state="disabled")
    attachment_path_entry.delete(0, "end")
    attachment_path_entry.config(state="disabled")
    attachment_name_entry.delete(0, "end")
    attachment_name_entry.config(state="disabled")
    send_attachment_button.config(state="disabled")
    send_jsutmsg_button.config(state="normal")

def toggle_attachment_buttons(choice):
    if choice == 1:
        one_button.config(bg="light blue", selectcolor="light blue")
        multiple_button.config(bg="#0563bb", selectcolor="#0563bb")
    elif choice == 2:
        one_button.config(bg="#0563bb", selectcolor="#0563bb")
        multiple_button.config(bg="light blue", selectcolor="light blue")
    browse_attachment_button.config(state="normal")

def check_attachment_choice(choice):      
    if choice == 1:
        browse_one_attachment()
    elif choice == 2:
        browse_multiple_attachment()
    else:
        browse_attachment_button.config(state="disabled")

def browse_one_attachment():
    filename = filedialog.askopenfilename(initialdir="C:/", title="Please Select the Attachment")
    attachment_path_entry.config(bg="white")
    attachment_path_entry.delete(0, "end")
    attachment_path_entry.insert(0, filename)

def browse_multiple_attachment():
    filename = filedialog.askdirectory(initialdir="C:/", title="Please Select the Attachment File Directory")
    attachment_path_entry.config(bg="white")
    attachment_path_entry.delete(0, "end")
    attachment_path_entry.insert(0, filename)

def browse_file():
    filename = filedialog.askopenfilename(initialdir="C:/", title="Please Select the Excel File")
    excel_file_entry.config(bg="white")
    excel_file_entry.delete(0, "end")
    excel_file_entry.insert(0, filename)

def browse_msg_text():
    filename = filedialog.askopenfilename(initialdir="C:/", title="Please Select the Message Template File")
    message_text_entry.config(bg="white")
    message_text_entry.delete(0, "end")
    message_text_entry.insert(0, filename)

def get_widgets_config(widgets_list):
    widgets_config_keys = []
    widgets_config_items = []
    for widget in widgets_list:
        single_widget_config_keys = []
        single_widget_config_items = []
        for key in widget.keys():
            single_widget_config_keys.append(key)
            single_widget_config_items.append(widget.cget(key))
        widgets_config_keys.append(single_widget_config_keys)
        widgets_config_items.append(single_widget_config_items)
    return widgets_config_keys, widgets_config_items

def reset_widgets_config(widgets_list, widgets_config_keys, widgets_config_items):
    i = 0
    for widget in widgets_list:
        if type(widget) == type(Entry()):
            widget.delete(0, "end")
        k = 0
        for key in widgets_config_keys[i]:
            widget[key] = widgets_config_items[i][k]
            k = k + 1
        i = i + 1

def reset_all(entries_list, entries_config_keys, entries_config_items, buttons_list, buttons_config_keys, buttons_config_items):
    app.focus()
    attachment_choice.set(0)
    browse_attachment_button.invoke()
    reset_widgets_config(entries_list, entries_config_keys, entries_config_items)
    reset_widgets_config(buttons_list, buttons_config_keys, buttons_config_items)
    app.update_idletasks()
    # i = 0
    # for entry in entries_list:
    #     entry.delete(0, "end")
    #     k = 0
    #     for key in entries_config_keys[i]:
    #         entry[key] = entries_config_items[i][k]
    #         k = k + 1
    #     i = i + 1
    # i = 0
    # for button in buttons_list:
    #     k = 0
    #     for key in buttons_config_keys[i]:
    #         button[key] = buttons_config_items[i][k]
    #         k = k + 1
    #     i = i + 1
    
# HEADER
heading = Label(text="تطبيق مساق لإرسال الملفّات", bg="#0563bb", fg="white", font="20", width="500", height="2")
heading.pack()
subheading = Label(text="Kindly Fill the Form Below", bg="white", fg="#0563bb", font="10")
subheading.pack()

# LABELS
email_sender = Label(text="Email Address", bg="white")
email_sender.place(x=15, y=90)
email_password = Label(text="Password", bg="white")
email_password.place(x=300, y=90)
excel_file = Label(text="Recipents Excel File", bg="white")
excel_file.place(x=15, y=140)
attachment_question = Label(text="Are You Sending an Attachment?", bg="white")
attachment_question.place(x=15, y=190)
attachment_quantity_question = Label(text="If YES, Select the Proper Option", bg="white")
attachment_quantity_question.place(x=295, y=190)
attachment_path = Label(text="Attachment File Path Or Directory", bg="white")
attachment_path.place(x=15, y=270)
message_text = Label(text="Message Template Path", bg="white")
message_text.place(x=15, y=320)
email_subject = Label(text="Email Subject", bg="white")
email_subject.place(x=15, y=370)
attachment_name = Label(text="Attachment Name", bg="white")
attachment_name.place(x=300, y=370)

# ENTRIES
temp_sender = StringVar()
temp_password = StringVar()
temp_receiver = StringVar()
temp_msg = StringVar()
temp_subject = StringVar()
temp_body = StringVar()
temp_attach_path = StringVar()
temp_attach_name = StringVar()
attachment_choice = IntVar() # for radio button
email_sender_entry = Entry(textvariable=temp_sender, width="40")
email_password_entry = Entry(textvariable=temp_password, width="40", show="•")
excel_file_entry = Entry(textvariable=temp_receiver, width="75", fg="green")
attachment_path_entry = Entry(textvariable=temp_attach_path, width="75", fg="green", state="disabled")
message_text_entry = Entry(textvariable=temp_msg, width="75", fg="green")
email_subject_entry = Entry(textvariable=temp_subject, width="40")
attachment_name_entry = Entry(textvariable=temp_attach_name, width="40", state="disabled")
email_sender_entry.place(x=15, y=110)
email_password_entry.place(x=300, y=110)
excel_file_entry.place(x=15, y=160)
attachment_path_entry.place(x=15, y=290)
message_text_entry.place(x=15, y=340)
email_subject_entry.place(x=15, y=390)
attachment_name_entry.place(x=300, y=390)
entries_list = [email_sender_entry, email_password_entry, excel_file_entry, attachment_path_entry, message_text_entry, email_subject_entry, attachment_name_entry]
entries_config_keys, entries_config_items = get_widgets_config(entries_list)
for entry in entries_list:
    entry.bind("<FocusIn>", handle_entry_focus_red)

# BUTTONS
yes_button = Button(text="YES", bg="#0563bb", fg="white", width="5", command=lambda:[change_yesbtn_col()])
yes_button.place(x= 90, y= 210)
no_button = Button(text="NO", bg="#0563bb", fg="white", width="5", command=lambda:[change_nobtn_col()])
no_button.place(x= 90, y= 240)
one_button = Radiobutton(text="Single Attachment to all Students", fg="white", bg="#0563bb", selectcolor="#0563bb", state="disabled", variable=attachment_choice, value=1, command=lambda:[toggle_attachment_buttons(attachment_choice.get())])
one_button.place(x= 300, y= 210)
multiple_button = Radiobutton(text="Different Attachment to each Student", fg="white", bg="#0563bb", selectcolor="#0563bb", state="disabled", variable=attachment_choice, value=2, command=lambda:[toggle_attachment_buttons(attachment_choice.get())])
multiple_button.place(x= 300, y= 240)
browse_excel_button = Button(text="Browse File", bg="#0563bb", fg="white", command=browse_file)
browse_excel_button.place(x=475, y=158)
browse_attachment_button = Button(text="Browse File", bg="#0563bb", fg="white", state="disabled", command=lambda:[check_attachment_choice(attachment_choice.get())])
browse_attachment_button.place(x=475, y=288)
browse_msg_button = Button(text="Browse File", bg="#0563bb", fg="white", command=browse_msg_text)
browse_msg_button.place(x=475, y=338)
send_jsutmsg_button = Button(text="Send Emails with No Attachments", bg="#0563bb", fg="white", width="35", height="2", state="disabled", command=lambda:[check_entries(entries_list), main.finisher_msg()])
send_jsutmsg_button.place(x=20, y=430)
send_attachment_button = Button(text="Send Emails with Attachments", bg="#0563bb", fg="white", width="35", height="2", state="disabled", command=lambda:[check_entries(entries_list), main.finisher_attachment()])
send_attachment_button.place(x=20, y=470)
reset_button = Button(text="Reset the Form", bg="#0563bb", fg="white", width="35", height="2", command=lambda:[reset_all(entries_list, entries_config_keys, entries_config_items, buttons_list, buttons_config_keys, buttons_config_items)])
reset_button.place(x=285, y=430)
buttons_list = [yes_button, no_button, one_button, multiple_button, browse_excel_button, browse_attachment_button, browse_msg_button, send_jsutmsg_button, send_attachment_button, reset_button]
buttons_config_keys, buttons_config_items = get_widgets_config(buttons_list)

# FOOTER
footer = Label(text="This Application was Developed by Ibrahim Maassarani, Abed-AlRahman Kamoun, and Issa Kassas", bg="white")
footer.place(x=20, y=550)

app.mainloop()