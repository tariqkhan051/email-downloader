import re
import datetime
import os
import shutil
import win32com.client
from pathlib import Path, PureWindowsPath
import win32wnet
import json


#load configurations
with open("config.json") as jsonFile:
    jsonObject = json.load(jsonFile)
    jsonFile.close()

allowed_extensions = jsonObject['allowed_extensions']
number_of_days_check = jsonObject['number_of_days_check']
ignore_emails_from_senders = jsonObject['ignore_emails_from_senders']
only_unread = jsonObject['only_unread']
current_employees = jsonObject['current_employees']
overwrite_or_skip_or_rename = jsonObject['overwrite_or_skip_or_rename']
localPath = jsonObject['localPath']
networkPath = jsonObject['networkPath']
currentEmployeesPath = jsonObject['currentEmployeesPath']
newEmployeesPath = jsonObject['newEmployeesPath']

#manual configurations

#allowed_extensions = ['.pdf', '.docx']
#number_of_days_check = 15
#ignore_emails_from_senders = ['noreply']
#only_unread = False
#current_employees = "CURRENT EMPLOYEES 2022"

# 1=overwrite, 2=skip, 3=rename
#overwrite_or_skip_or_rename = 3

# example: r"E:/scripts/py/email_downloader/HR/"
#localPath = r"temp/"

# example1: r"Z:/scripts/py/email_downloader/HR/"
# example2: \\192.168.1.30\Shared\HR
#networkPath = r"//192.168.1.30/Shared/HR/"

# example: r"E:/scripts/py/email_downloader/HR/current/"
currentEmployeesPath = r"temp/"+current_employees+"/"

#example: r"E:/scripts/py/email_downloader/HR/new/"
#newEmployeesPath = r"temp/new/"

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
Filter = ("@SQL=" + chr(34) + "urn:schemas:httpmail:hasattachment" +
          chr(34) + "=1")

def ignore_sender(email):
    return email.lower().startswith(tuple(ignore_emails_from_senders))

def is_date_in_range(receivedDate):
    today = datetime.date.today()
    if number_of_days_check <= 1:
        return receivedDate == today
    else:
        daysCount = abs((today - receivedDate).days)
        return daysCount <= number_of_days_check

def is_valid_message(message):
    if (message.Class != 43):
        return False
    if (only_unread == True):
        if (message.Unread == False):
            return False
    print ("Is valid message!")
    return True

def is_message_older(message):
    if is_date_in_range(message.Senton.date()) == False:
        return True
    return False
    
def is_document(fileName):
    file = os.path.splitext(fileName)
    if file != None:
        print ("File is not none.")
        file_extension = file[1]
        print ("File extension:  " + file_extension)
        if allowed_extensions.count(str(file_extension).lower()) > 0:
            print ("Is Valid Extension")
            return True
    else:
        print ("File is none.")
        if fileName.lower().endswith(tuple(allowed_extensions)):
            print ("file has valid extensions")
            return True
    print ("\nIs not valid document!")
    return False
    
def copy_to_local_drive():
    host = "192.168.1.30"
    dest_share_path = "\\Shared\\HR"
    username = "irfan"
    password = "chickenroll1"
    toDir = localPath
    fromDir = networkPath
    print ("Local Drive Copying...")
    win32wnet.WNetAddConnection2(0, None, '\\\\'+host, None, username, password)
    #shutil.copy(source_file, '\\\\'+host+dest_share_path+'\\')
    destination = shutil.copytree('\\\\'+host+dest_share_path+'\\', toDir, dirs_exist_ok=True)
    win32wnet.WNetCancelConnection2('\\\\'+host, 0, 0) # optional disconnect
    print ("Connection Disconnected.")
    
    #if os.path.exists(toDir):
    #    shutil.rmtree(toDir)
    #destination = shutil.copytree(fromDir, toDir, dirs_exist_ok=True)

def get_files_tree(src):
    print ("Get Source Files Begin")
    req_files = []
    for r, d, files in os.walk(src):
        for file in files:
            src_file = os.path.join(r, file)
            src_file = src_file.replace('\\', '/')
            print ("Source File : " + str(src_file))
            if src_file.lower().endswith(tuple(allowed_extensions)):
                req_files.append(src_file)
    return req_files

def copy_tree(src_path, dest_path):
    print ("Copy Tree Begins")
    print ("src path : " + src_path)
    print ("dest path : " +dest_path)
    for cf in get_files_tree(src_path):
        df= cf.replace(src_path, dest_path)
        print ("dest file name : " + df)
        dest_dir = os.path.dirname(df)
        print ("dest directory: " + dest_dir)
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir)
            print ("directory created")
        else:
            print ("path exists")
        
        print ("Check Dest Dir : " + df)
        if os.path.exists(df):
            print ("File path exists.")
            if  overwrite_or_skip_or_rename == 1:
                print ("File to be replaced.")
                shutil.copy2(cf, df)
            elif  overwrite_or_skip_or_rename == 2:
                print ("File to be ignored.")
            elif  overwrite_or_skip_or_rename == 3:
                print ("File to be renamed and copied.")
                file = os.path.splitext(df)
                file_name = file[0]
                new_file_name = file_name + "_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                file_extension =  file[1]
                new_df = new_file_name + file_extension
                print ("new file path : " + str(new_df))
                shutil.copy2(cf, new_df)
        else:
            shutil.copy2(cf, df)
            print ("File Copied.")
    
def copy_to_network_drive():
    destination = shutil.copytree(src, dest, dirs_exist_ok=True)

def is_current_employee_via_dirNme(sender_name, sender_email):
    names = [name for name in os.listdir(currentEmployeesPath) if os.path.isdir(os.path.join(currentEmployeesPath,name))]
    for name in names:
        formattedName = name.lower().strip()
        if (formattedName == sender_email.lower().strip() or formattedName == sender_name.lower().strip()):
            return [True, name]
    return [False, '']
    
def is_current_employee(sender_email):
    path = Path("current_employees_email_addresses.txt")
    if (path.is_file()):
        namesFile = open(path, 'r')
        names = namesFile.readlines()
        for name in names:
            if (name.lower().strip() == sender_email.lower().strip()):
                return True
    return False
    
try:
    messages = inbox.Items.Restrict(Filter)
    messages.Sort('[ReceivedTime]', True)
    copied_to_local = False
    for message in messages:
        if (is_message_older(message)):
            print ("Last Date to Ignore From : ")
            print (message.Senton.date())
            break
        if is_valid_message(message):
            if message.SenderEmailType=='EX':
                if message.Sender.GetExchangeUser() != None:
                    current_sender = str(message.Sender.GetExchangeUser().PrimarySmtpAddress).lower()
                else:
                    current_sender = str(message.Sender.GetExchangeDistributionList().PrimarySmtpAddress).lower()
            else:
                current_sender = str(message.SenderEmailAddress).lower()
            
            if ignore_sender(current_sender):
                print ("Sender Ignored: ")
                print (current_sender)
                continue
                
            current_subject = str(message.Subject).lower()
            current_sender_name = str(message.Sender).lower()
            for attachment in message.Attachments:
                fileName = attachment.FileName
                if (is_document(fileName)):
                    if (copied_to_local == False):
                        copied_to_local = True
                        copy_to_local_drive()
                    
                    checkSender = is_current_employee_via_dirNme(current_sender_name, current_sender)
                    if (checkSender[0]):
                        newPath = currentEmployeesPath + checkSender[1]
                    else:
                        print ("New path to be created.")
                        newPath = newEmployeesPath + current_sender
                    Path(newPath).mkdir(parents=True, exist_ok=True)
                    fname = os.path.join(newPath, str(attachment))
                    current_dir = os.getcwd()
                    print (current_dir)
                    filename = PureWindowsPath(current_dir + "/" + fname)
                    correct_path = Path(filename)
                    if (os.path.isfile(correct_path) == False):
                        print ("Saving File: " + str(correct_path))
                        attachment.SaveAsFile(correct_path)
        else:
            pass
    copy_tree(localPath, networkPath)
except Exception as e:
    print('Error: '+ str(e))
    
# while message != None and emailCount > 0:
  # try:
    
    # attachments = message.Attachments
    # attachmentCount = len(attachments)
    # print("Attachment Count: " + str(attachmentCount))
    # if attachmentCount > 0:
        # for attachment in message.Attachments:
            # print(attachment.FileName)
            # attachment.SaveAsFile(os.path.join(path, str(attachment)))
                
        # attachment = attachments.Item(1)
        # attachment_name = str(attachment).lower()
        # attachment.SaveASFile(path + '\\' + attachment_name)
    # else:
        # pass
    # message = messages.GetNext()
    # emailCount -= 1
  # except Exception as e:
    # print('Error: '+ str(e))
    # message = messages.GetNext()
    # emailCount -= 1
# exit
