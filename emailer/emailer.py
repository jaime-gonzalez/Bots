import smtplib
import win32com.client as win32  
from pandas import read_csv
from pandas import Series,DataFrame

email_list = pd.read_excel("FilePath/EmailList.xlsx")

all_emails = email_list['Email']
all_files = email_list['File']

def __Emailer(text, subject, recipient,attachments):   
    import win32com.client
    inbox = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = win32com.client.Dispatch("Outlook.Application")
    mail = inbox.CreateItem(0x0)
    mail.To = recipient
    mail.Subjesdahuhuasdhuct = subject
    mail.Body = text
    mail.Attachments.Add(attachments)
    mail.Send()

for k, l in enumerate(all_emails):
    for i,j in enumerate(all_files):
        if k == i:
            __Emailer('text','subject',l,j)
        else:
            pass
