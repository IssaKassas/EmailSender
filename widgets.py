from tkinter import Button, Entry, Frame, IntVar, Label, StringVar, Tk , Radiobutton
from connection import send
from methods import browse_excel_file, browse_path_files , browse_template_file
from methods import get_widget_config, handle_button_focus_red, handle_entry_focus_red
from methods import noButton, yesButton, toggleRadioButton , check_entries

window = Tk()

window.title("Email Sender Application")

window.geometry("800x650")

window.config(bg = "white")

window.resizable(False , False)

headerFrame = Frame(
    window , 
    bg = "white"
)

inputFrame = Frame(
    window,
    bg = "white"
)

questionFrame = Frame(
    window,
    bg = "white"
)

pathFrame = Frame(
    window,
    bg = "white"
)

detailsFrame = Frame(
    window,
    bg = "white"
)

buttonsFrame = Frame(
    window,
    bg = "white"
)

footerFrame = Frame(
    window,
    bg = "white"
)

header = Label(
    headerFrame,
    text = "Email Sender Application",
    bg = "white",
    font = ('Arial 15 bold')
)

subheader = Label(
    headerFrame,
    text = "Kindly fill the form bellow",
    bg = "white",
    font= "12"
)

emailLabel = Label(
    inputFrame,
    text = "Email Address",
    bg = "white"
)

email = StringVar()
emailEntry = Entry(
    inputFrame,
    width = "30",
    textvariable = email
)

passwordLabel = Label(
    inputFrame,
    text = "Password",
    bg = "white"
)

password = StringVar()
passwordEntry = Entry(
    inputFrame, 
    width = "25",
    show = "*",
    textvariable = password
)

excelLabel = Label(
    inputFrame,
    text = "Receiver Excel File",
    bg = "white"
)

excelFile = StringVar()
excelEntry = Entry(
    inputFrame,
    width = "50",
    fg = "green",
    textvariable = excelFile  
)

browseExcelFile = Button(
    inputFrame,
    text= "Browse Excel File",
    bg = "black",
    fg = "white",
    command = lambda : browse_excel_file(excelEntry)
)

attachmentLabel = Label(
    questionFrame,
    text = "Are you sending an Antachment?",
    bg = "white"
)

yes = Button(
    questionFrame,
    text = "YES",
    bg = "black",
    fg = "white",
    width = "8",
    command= lambda: yesButton(attachment_entry , filename ,singleRadio , multipleRadio , 
                               yes , no , browse_attachment , sendEmail , deleteEntries
                               )
)

no = Button(
    questionFrame,
    text = "NO",
    bg = "black",
    fg = "white",
    width = "8",
    command= lambda: noButton(attachment_entry , filename ,singleRadio , multipleRadio , 
                               yes , no , browse_attachment , sendEmail , deleteEntries
                               )
)

optionLabel = Label(
    questionFrame,
    text = "if YES, Select the proper option",
    bg = "white"
)

choice = IntVar()

singleRadio = Radiobutton(
    questionFrame,
    fg = "black",
    text = "Single Attachment to all students",
    bg = "white",
    selectcolor= "white",
    value = 1,
    state = "disabled",
    variable = choice,
    command = lambda: toggleRadioButton(choice , singleRadio , multipleRadio , browse_attachment)
)

multipleRadio = Radiobutton(
    questionFrame,
    text = "Different Attachment to each Student",
    bg = "white",
    fg = "black",
    value = 2,
    selectcolor= "white",
    state = "disabled",
    variable = choice,
    command = lambda: toggleRadioButton(choice , singleRadio , multipleRadio , browse_attachment)
)

attachment_path = Label(
    pathFrame,
    bg = "white",
    text = "Attachment File path or Directory",
)

attachment = StringVar()
attachment_entry = Entry(
    pathFrame, 
    width = "75",
    state = "disabled",
    textvariable = attachment
)

browse_attachment = Button(
    pathFrame,
    text = "Browse Attachment File",
    bg = "black",
    fg= "white",
    state = "disabled",
    command = lambda: browse_path_files(choice , browse_attachment , attachment_entry)
)

message_template = Label(
    detailsFrame,
    text  = "Message Template Path",
    bg  = 'white'
)

templateFile = StringVar()
message_entry = Entry(
    detailsFrame,
    width = "50",
    fg= "green",
    textvariable = templateFile
)

browse_message = Button(
    detailsFrame,
    text = "Browse Template File",
    bg = "black",
    fg= "white",
    command = lambda: browse_template_file(message_entry)
)

subject_label = Label(
    detailsFrame,
    text="Email Subject",
    bg = "white"
)

subject = StringVar()
subject_entry = Entry(
    detailsFrame,
    width = "40",
    textvariable = subject
)

fileLabel = Label(
    detailsFrame,
    text = "Attachment Name",
    bg = "white"
)

attachment_name = StringVar()
filename = Entry(
    detailsFrame,
    width = "40",
    state = "disabled",
    textvariable = attachment_name
)

deleteEntries = [attachment_entry , message_entry , subject_entry , filename]

entriesList = [emailEntry , passwordEntry , excelEntry , attachment_entry ,
                message_entry ,subject_entry , filename
               ]

entries_config_keys, entries_config_items = get_widget_config(entriesList)

for entry in entriesList:
    entry.bind("<FocusIn>", handle_entry_focus_red)

sendEmail = Button(
    buttonsFrame,
    text = "Send",
    width= "25",
    bg = "black",
    fg="white",
    height = 3,
    state = "disabled",
    command = lambda: [check_entries(window , entriesList), 
                send(email , password , templateFile , subject , excelFile ,
                           attachment , attachment_name , attachment_entry)]
)

reset = Button(
    buttonsFrame,
    text = "Reset the Form",
    width= "25",
    bg = "black",
    fg="white",
    height = 3
)

Exit = Button(
    buttonsFrame,
    text = "Exit",
    width= "25",
    bg = "black",
    fg="white",
    height = 3,
    command=window.destroy
)

buttonsList = [yes , no , browse_attachment , browse_message ,
                browseExcelFile , sendEmail , reset , Exit
              ]

buttons_config_keys, buttons_config_items = get_widget_config(buttonsList)

for btn in buttonsList:
    btn.bind("<FocusIn>", handle_button_focus_red)

footer = Label(
    footerFrame,
    text = "This Application was developped by Abdelghani Shaaban & Mohamad Laalaa",
    bg = "white"
)