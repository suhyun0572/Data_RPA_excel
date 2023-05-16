import win32com.client as cli
from datetime import datetime
import os
import time
import win32com.client

# 3_2. edit_print
from edit_print import edit_print
# 3_3. mail_sender
from mail_sender import mail_sender


def download_schedules():
    # sender names are saved in a text file who send me the files.
    f = open("sendername.txt", 'r')

    # read and make a sender list
    senders = f.readlines()
    senderlist = []
    for line in senders:
        senderlist.append(line[:-1])
    f.close()

    # read and make a sender email list
    f_mail = open("sendermail.txt", 'r')
    senders_mail = f_mail.readlines()
    sender_mail_list = []
    for line in senders_mail:
        sender_mail_list.append(line[:-1])
    f.close()

    # a list which will be checked and found
    senderlist+=sender_mail_list

    # path
    download_path = r" ... "
    save_path = r" ... "
    sending_file_list = []
    outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI")
    box = outlook.GetDefaultFolder(6)
    mes = box.Items
    for mail in mes:
        for sender in senderlist:
            if sender in mail.SenderName or sender in mail.SenderEmailAddress:
                attachments = mail.Attachments
                for i in range(1,attachments.count+1) :
                    attachment = attachments.Item(i)
                    if str(attachment)[-5:] == ".xlsx":
                        if str(attachment) in os.listdir(download_path):
                            pass
                        else :
                            # save the attached files
                            attachment_count = len(os.listdir(download_path))+1
                            attachment.SaveASFile(download_path + "[" + str(attachment_count)+ "] " +str(mail.ReceivedTime).split(' ')[0].replace('-','_')+".xlsx")

                            # analyze and save the analyzed file with "edit_print"
                            edit_print()

                            # a list which is the analyzed and new created excel file 
                            sending_file_list.append(save_path+"[" + str(attachment_count)+ "] " +str(mail.ReceivedTime).split(' ')[0].replace('-','_')+'_updated'+".xlsx")
    # Delete the received email that has been used.                        
    for mail_del in mes :
        if " ... " in mail_del.Subject :
            mail_del.Delete()

    # send to charger the analyzed and updated file
    mail_sender(sending_file_list)

if __name__=="__main__":
    download_schedules()
