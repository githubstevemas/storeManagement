import os
from datetime import datetime, timedelta
import shutil

ROOT_FOLDER = "saves"


def get_dates():
    today = datetime.today()
    if today.weekday() == 0:  # lundi
        prev_day = today - timedelta(days=3)  # vendredi précédent
    else:
        prev_day = today - timedelta(days=1)
    return today, prev_day


def create_folder(folder_path):
    os.makedirs(ROOT_FOLDER, exist_ok=True)
    os.makedirs(folder_path, exist_ok=True)
    print(f"Dossier créé ou existant : {folder_path}")


def move_pdf(target_folder, folder='.'):
    for fichier in os.listdir(folder):
        if fichier.lower().endswith('.pdf'):
            shutil.move(fichier, os.path.join(target_folder, fichier))
            print(f"{fichier} déplacé dans {target_folder}")
