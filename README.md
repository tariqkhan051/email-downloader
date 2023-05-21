# Email downloader
this script helps downloading specified files from logged in email account on PC and move to shared drives

# How it works

-	Based on the configurations defined in ```config.json``` email attachments will be downloaded
-	Folders in Network Drive (Example: Z:/HR/Current) will be used to find current employees names, whereas folder names could be employees’ names or their email address
![image](https://github.com/tariqkhan051/email-downloader/assets/15242136/8a9825e9-b328-4fc3-bf9b-aeaf9bc16bdf)
-	All directories from the defined Network drive path will be first copied to local drive
-	If the sender email or sender name matches with the list of names we collected in previous step, then the document will be downloaded in the *currentEmployeesPath*, otherwise it will be downloaded in the *newEmployeesPath*
-	Folders on *localPath*  will be copied to Network Drive path i.e. *networkPath*

# Prerequisites
[Python](https://www.python.org/downloads/) installed on your PC.

Outlook App configured with your email account.

# Steps to run
Update configurations in ``` config.json ``` file.

## Configurations

| Parameter | Description | Example
| --- | --- | --- |
| allowed_extensions | Emails with attachments having defined extensions will be considered | ['.pdf', '.docx'] |
| number_of_days_check | Days to consider for emails. If 1, then only today’s emails will be considered | 2 |
| ignore_emails_from_senders | Emails received from the sender email addresses starting from the defined list | ['noreply'] |
| only_unread | If sets to True, then only unread emails will be considered | True/False |
| localPath | Path on local pc drive to copy directories from network drive | |
| networkPath | Network drive to copy directories from local drive | |
| currentEmployeesPath | Current Employees Folder path inside localPath | |
| newEmployeesPath | New Employees Folder path inside localPath | | 

## Run in CMD

Run ``` python read_email.py ``` to start the program


