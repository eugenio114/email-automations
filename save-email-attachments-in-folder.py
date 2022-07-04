import win32com.client
import os
import datetime as dt

desktop = os.path.expanduser('~') + '/Desktop/'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print (dt.datetime.now())

# setup range for outlook to search emails (so we don't go through the entire inbox)
lastQuarterDateTime = dt.datetime.now() - dt.timedelta(days = 90)
lastQuarterDateTime = lastQuarterDateTime.strftime('%m/%d/%Y %H:%M %p')

# Select main Inbox
inbox = outlook.GetDefaultFolder(6)

# Optional:  Select main Inbox, look in subfolder "Test"
#inbox = outlook.GetDefaultFolder(6).Folders["Test"]

messages = inbox.Items

# Only search emails in the time range above:
messages = messages.Restrict("[ReceivedTime] >= '" + lastQuarterDateTime +"'")

print ('Reading Inbox, including Inbox Subfolders...')

# Download a select attachment ---------------------------------------
# Create a folder to capture attachments.
outlook_export_folder = f'{desktop}Outlook Export/'
if not os.path.exists(outlook_export_folder): os.makedirs(outlook_export_folder)

try:
    for message in list(messages):
        try:
            s = message.sender
            s = str(s)
            print('Sender:' , message.sender)
            for att in message.Attachments:
                # Give each attachment a path and filename
                outfile_name1 = outlook_export_folder + att.FileName
                # save file
                att.SaveASFile(outfile_name1)
                print('Saved file:', outfile_name1)

        except Exception as e:
            print(f"type error: {str(e)}")
            x=1

except Exception as e:
    print(f"type error: {str(e)}")
    x=1

#move files into dedicated folders depending on file type (like .png / .jpg)-----------------------------------------

main_folder = os.listdir(outlook_export_folder)
image_folder = f'{outlook_export_folder}IMAGES/'
if not os.path.exists(image_folder): os.makedirs(image_folder)
ppt_folder = f'{outlook_export_folder}PPTs/'
if not os.path.exists(ppt_folder): os.makedirs(ppt_folder)
documents_folder = f'{outlook_export_folder}DOCUMENTS/'
if not os.path.exists(documents_folder): os.makedirs(documents_folder)
email_folder = f'{outlook_export_folder}EMAILS/'
if not os.path.exists(email_folder): os.makedirs(email_folder)
excel_folder = f'{outlook_export_folder}EXCEL-CSV/'
if not os.path.exists(excel_folder): os.makedirs(excel_folder)


for item in main_folder:
    if item.endswith(".png") or item.endswith(".jpg"):
        os.rename(os.path.join(outlook_export_folder, item), (os.path.join(outlook_export_folder,image_folder,item)))

for item in main_folder:
    if item.endswith(".pptx"):
        os.rename(os.path.join(outlook_export_folder, item), (os.path.join(outlook_export_folder,ppt_folder,item)))

for item in main_folder:
    if item.endswith(".pdf") or item.endswith(".docx"):
        os.rename(os.path.join(outlook_export_folder, item), (os.path.join(outlook_export_folder,documents_folder,item)))

for item in main_folder:
    if item.endswith(".msg"):
        os.rename(os.path.join(outlook_export_folder, item), (os.path.join(outlook_export_folder,email_folder,item)))

for item in main_folder:
    if item.endswith(".xlsx") or item.endswith(".csv") or item.endswith(".xls"):
        os.rename(os.path.join(outlook_export_folder, item), (os.path.join(outlook_export_folder,excel_folder,item)))