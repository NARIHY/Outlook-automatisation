import win32com.client
import re
from datetime import datetime

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#  Dossier 'Projets' existant
inbox = outlook.GetDefaultFolder(6)  # Boîte de réception
projets_folder = inbox.Folders["projet"]  # À adapter si placé ailleurs

messages = projets_folder.Items
messages.Sort("[ReceivedTime]", True)

def safe_name(name):
    return name.replace(":", "-").replace("/", "-").replace("\\", "-")

def extract_project_code(text):
    match = re.search(r"T\d{3,5}-[A-Z0-9\-]+", text)
    return match.group(0) if match else None

def get_or_create_folder(parent, name):
    for f in parent.Folders:
        if f.Name == name:
            return f
    return parent.Folders.Add(name)

for mail in list(messages):  # Important de forcer la conversion, sinon bug de collection modifiée
    try:
        subject = mail.Subject or ""
        body = mail.Body or ""
        received = mail.ReceivedTime
        date_str = received.strftime("%Y-%m-%d")
        time_str = received.strftime("%Hh%M")

        project_code = extract_project_code(subject + " " + body)
        if not project_code:
            continue

        project_code = safe_name(project_code)

        #  Re-crée la structure dans le dossier Projets
        project_folder = get_or_create_folder(projets_folder, project_code)
        date_folder = get_or_create_folder(project_folder, date_str)
        time_folder = get_or_create_folder(date_folder, time_str)

        # 📤 Déplace le mail
        mail.Move(time_folder)

    except Exception as e:
        print("Erreur:", e)
