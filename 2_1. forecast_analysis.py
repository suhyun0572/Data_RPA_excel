import win32com.client as cli
from datetime import datetime
import os
from forcast_maker_v2 import edit_maker

# path
download_path = r" ... "
outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI")
box = outlook.GetDefaultFolder(6)
mes = box.Items

for mail in mes:
    # find ... in the email subject
    if " ... " in mail.Subject:
        attachments = mail.Attachments
        name_list = os.listdir()
        # Find and download the attached file.
        for i in range(1,attachments.count+1) :
            attachment = attachments.Item(i)
            if str(attachment) in name_list :
                pass
            else:
                if str(attachment)[-5:] == '.xlsx':
                    attachment.SaveASFile(download_path + str(attachment))

# data analysis and save in another folder                    
edit_maker()