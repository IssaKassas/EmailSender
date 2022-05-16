from widgets import *
from tkinter import W

header.grid(row = 0 , column = 0 , pady = 10)

subheader.grid(row = 1 , column = 0)

headerFrame.grid(row = 0 , column = 0 , pady = 20)

emailLabel.grid(row = 0 , column = 0 , sticky = W , padx = 100)

emailEntry.grid(row = 1 , column = 0 , pady = 10 , padx= 100)

passwordLabel.grid(row = 0 , column= 1, padx = 100)

passwordEntry.grid(row = 1 , column= 1 , padx = 100)

excelLabel.grid(row= 2 , column = 0 , sticky = W  , padx = 40, pady = 15)

excelEntry.grid(row = 3, column = 0)

browseExcelFile.grid(row = 3 , column = 1, padx = 10)

inputFrame.grid(row = 1 , column = 0, padx= 10 , sticky= W)

attachmentLabel.grid(row = 0 , columnspan = 2)

yes.grid(row= 1, column = 0 , pady = 6)

no.grid(row = 1 , column = 1)

optionLabel.grid(row = 0 , column = 2, columnspan = 2,  pady = 5, padx= 20)

singleRadio.grid(row = 1 , column= 2 , pady = 5 , padx= 70)

multipleRadio.grid(row = 1 , column = 3)

questionFrame.grid(row = 2 , column = 0 , padx=10 , pady = 30)

attachment_path.grid(row =  0, column = 0, sticky= W , padx = 20)

attachment_entry.grid(row = 1, column = 0, pady = 5 , padx = 20)

browse_attachment.grid(row = 1 , column = 1 , padx = 20)

pathFrame.grid(row = 3 , column = 0 , sticky = W , padx = 10)

message_template.grid(row = 0 , column= 0 , sticky = W)

message_entry.grid(row = 1 , column = 0)

browse_message.grid(row = 1 , column = 1 , padx = 10)

subject_label.grid(row = 2 , column = 0 , sticky = W)

subject_entry.grid(row = 3 , column= 0 , sticky = W)

fileLabel.grid(row = 2 , column= 1 , pady= 5)

filename.grid(row = 3 , column = 1)

detailsFrame.grid(row = 4 , column= 0 , sticky=W , padx  = 50 , pady = 10)

sendEmail.grid(row = 0, column=0)

reset.grid(row = 0 , column = 1 , padx = 10)

Exit.grid(row = 0 , column= 2 , padx = 10)

buttonsFrame.grid(row = 5 , column = 0 , pady = 10)

footer.grid(row = 0 , column = 0)

footerFrame.grid(row = 6, columnspan = 1)

window.mainloop()