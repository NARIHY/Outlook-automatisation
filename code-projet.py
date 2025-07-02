import win32com.client
import re
from datetime import datetime

# Initialisation Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

# Dictionnaire mois en français
MONTH_NAMES = {
    1: "janvier", 2: "février", 3: "mars", 4: "avril",
    5: "mai", 6: "juin", 7: "juillet", 8: "août",
    9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre"
}

def safe_name(name):
    """Remplace les caractères invalides pour le nom de dossier."""
    return name.replace(":", "-").replace("/", "-").replace("\\", "-")

def extract_project_code(text):
    """
    Extrait un code projet au format T###-XXXX-... depuis le sujet ou le corps.
    Exemple : T1112-BTOC-IFOP-QPV
    """
    match = re.search(r"T\d{3,5}-[A-Z0-9\-]+", text)
    return match.group(0) if match else None


def get_or_create_folder(parent, folder_name):
    """Retourne le sous-dossier existant ou le crée s’il n’existe pas."""
    for f in parent.Folders:
        if f.Name == folder_name:
            return f
    return parent.Folders.Add(folder_name)

# Création du dossier racine "Projets"
projets_root = get_or_create_folder(inbox, "Projets")

for mail in messages:
    try:
        subj = mail.Subject or ""
        body = mail.Body or ""
        received = mail.ReceivedTime  # datetime

        # Extrait le code projet
        code = extract_project_code(subj + " " + body)
        if not code:
            continue  # on ignore les mails sans code projet

        code = safe_name(code)

        # Année
        year_str = str(received.year)

        # Nom du mois en toutes lettres
        month_name = MONTH_NAMES.get(received.month, str(received.month))

        # Arborescence : Projets → Année → Mois → CODE
        year_folder = get_or_create_folder(projets_root, year_str)
        month_folder = get_or_create_folder(year_folder, month_name)
        project_folder = get_or_create_folder(month_folder, code)

        # Déplace le mail
        mail.Move(project_folder)
        print(f"Déplacé : '{subj}' → Projets\\{year_str}\\{month_name}\\{code}")

    except Exception as e:
        print(f"Erreur pour '{mail.Subject}': {e}")
