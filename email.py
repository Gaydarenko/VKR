# install pywin32
import win32com.client
import os
import datetime as dt


# инициирование сеанса
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items

# осуществление фильтрации
received_point = dt.datetime.now() - dt.timedelta(days=1)
received_point = received_point.strftime("%m/%d/%Y %H:%M %p")
messages = messages.Restrict("[ReceivedTime] >= '" + received_point + "'")
messages = messages.Restrict("[SenderEmailAddress] = 'azazelik@mail.ru'")

print('0')
# os.mkdir('email_files')
current_dir = os.getcwd()
output_dir = 'email_files'
try:
    print(f'1 - {list(messages)}')
    for message in list(messages):
        try:
            s = message.sender
            for attachment in message.Attachments:
                print('5')
                attachment.SaveAsFile(os.path.join(current_dir, output_dir, attachment.FileName))
                print(f"attachement {attachment.FileName} from {s} saved")
        except Exception as e:
            print("error when saving the attachment:" + str(e))
except Exception:
    print("Ooops!!!")
