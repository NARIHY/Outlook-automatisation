import win32com.client

def move_items_to_inbox(folder, inbox_folder):
    """
    Parcourt récursivement `folder` et déplace tous les éléments qu'il contient
    dans `inbox_folder`.
    """
    # Copier la liste des items pour éviter les problèmes lors du déplacement
    items = list(folder.Items)
    for item in items:
        try:
            item.Move(inbox_folder)
            print(f"Déplacé : {item.Subject}")
        except Exception as e:
            print(f"Erreur lors du déplacement de l'élément '{item.Subject}': {e}")

    # Parcours des sous-dossiers
    for sub in folder.Folders:
        move_items_to_inbox(sub, inbox_folder)


def main():
    # Initialiser Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Récupérer le dossier Inbox par défaut (numéro 6)
    inbox = outlook.GetDefaultFolder(6)

    # Pour parcourir uniquement les sous-dossiers de la boîte de réception :
    for folder in inbox.Folders:
        print(f"Traitement du dossier : {folder.Name}")
        move_items_to_inbox(folder, inbox)

    print("Terminé : tous les mails des sous-dossiers ont été remis dans la boîte de réception.")

if __name__ == "__main__":
    main()
