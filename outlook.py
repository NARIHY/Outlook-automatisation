import win32com.client
from datetime import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

def safe_name(name):
    return name.replace(":", "-").replace("/", "-").replace("\\", "-")

for mail in messages:
    try:
        sender = mail.SenderName
        received = mail.ReceivedTime
        date_str = received.strftime("%Y-%m-%d")
        time_str = received.strftime("%Hh%M")

        # Noms sûrs pour les dossiers
        sender_folder = safe_name(sender)
        date_folder = safe_name(date_str)
        time_folder = safe_name(time_str)

        # Crée dossier expéditeur
        sender_folder_obj = inbox.Folders.GetFirst()
        if sender_folder not in [f.Name for f in inbox.Folders]:
            sender_folder_obj = inbox.Folders.Add(sender_folder)
        else:
            sender_folder_obj = inbox.Folders[sender_folder]

        # Crée dossier date
        if date_folder not in [f.Name for f in sender_folder_obj.Folders]:
            date_folder_obj = sender_folder_obj.Folders.Add(date_folder)
        else:
            date_folder_obj = sender_folder_obj.Folders[date_folder]

        # Crée dossier heure
        if time_folder not in [f.Name for f in date_folder_obj.Folders]:
            time_folder_obj = date_folder_obj.Folders.Add(time_folder)
        else:
            time_folder_obj = date_folder_obj.Folders[time_folder]

        # Déplacement du mail
        mail.Move(time_folder_obj)

    except Exception as e:
        print(f"Erreur avec un message : {e}")
