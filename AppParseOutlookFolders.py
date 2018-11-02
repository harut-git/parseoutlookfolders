import errno
import json
import os

import win32com.client

# -- Static variables -- #
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).parent


def parse_folder(folder_obj, mail_quantity):
    messages = folder_obj.Items
    message = messages.GetLast()
    i = 0
    output_list = []
    while message:
        if message.Class == 43:
            subject = message.Subject
            if message.SenderEmailType == "EX":
                sender = message.Sender.GetExchangeUser().PrimarySmtpAddress
            else:
                sender = message.SenderEmailAddress
            recipients = message.Recipients
            body = message.Body
            current_recipients = []
            for r in recipients:
                try:
                    k = r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                    current_recipients.append(k)
                except AttributeError:
                    k = r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
                    current_recipients.append(k)
            output_list.append({
                'subject': subject.encode("utf-8"),
                'from': sender,
                'recipients': current_recipients,
                'body': body.encode("utf-8")
            })
        i += 1

        message = messages.GetPrevious()
        if i > mail_quantity:
            break
    path = str(folder_obj.FolderPath).replace('\\', '/') + '.json'
    if path[:2] == '//':
        path = path[2:]
    final_path = 'Results/' + path
    if not os.path.exists(os.path.dirname(final_path)):
        try:
            print final_path
            os.makedirs(os.path.dirname(final_path))
        except OSError as exc:  # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise
    with open(final_path, "w+") as outfile:
        outfile.write(json.dumps(output_list))


def dig_folders(folder):
    parse_folder(folder, 100)
    if folder.Folders.Count > 0:
        for i in range(1, folder.Folders.Count + 1):
            dig_folders(folder.Folders(i))


dig_folders(inbox)
print "Completed Successfully..."
