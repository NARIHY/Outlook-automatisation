import win32com.client
import re
from datetime import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

def safe_name(name):
    return name.replace(":", "-").replace("/", "-").replace("\\", "-")

def extract_project_code(text):
    match = re.search(r"T\d{3,5}-[A-Z0-9\-]+", text)
    return match.group(0) if match else None

# Vérifie ou crée un sous-dossier
def get_or_create_folder(parent, folder_name):
    for f in parent.Folders:
        if f.Name == folder_name:
            return f
    return parent.Folders.Add(folder_name)

# 🔁 Parcours des emails
for mail in messages:
    try:
        subject = mail.Subject or ""
        body = mail.Body or ""
        received = mail.ReceivedTime
        date_str = received.strftime("%Y-%m-%d")
        time_str = received.strftime("%Hh%M")

        # Recherche code projet
        project_code = extract_project_code(subject + " " + body)
        if not project_code:
            continue  # skip les emails sans code projet

        #  Nettoyage
        project_code = safe_name(project_code)

        #  Structure de dossiers
        projets_folder = get_or_create_folder(inbox, "Projets")
        projet_folder = get_or_create_folder(projets_folder, project_code)
        date_folder = get_or_create_folder(projet_folder, date_str)
        time_folder = get_or_create_folder(date_folder, time_str)

        #  Déplacement
        mail.Move(time_folder)

    except Exception as e:
        print(f"Erreur pour un message : {e}")
