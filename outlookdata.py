from win32com.client import Dispatch
import csv
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
#Get Inbox Folder
inbox = outlook.GetDefaultFolder("6")
all_inbox = inbox.Items
folders = inbox.Folders
with open('large.csv','w') as f1:
    for msg in all_inbox:
    # Only Mail Items Objects from Outlook
        if msg.Class==43:
            try: 
                #Check for Internal Email addresses
                if msg.SenderEmailType=='EX':
                    writer= csv.writer(f1, delimiter='\t',lineterminator='\n',)
                    writer.writerow(msg.SenderEmailAddress)
                    print (msg.SenderEmailAddress)
            except: 
                pass