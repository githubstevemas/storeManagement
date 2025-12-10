import os
import shutil

from prepaProfil.upload.utils.folder import get_dates, create_folder, move_pdf
from prepaProfil.upload.utils.pdf_processing import extract_pdf_data, find_pdf


def main():
    today, prev_day = get_dates()
    today_str = today.strftime('%d-%m-%Y')
    prev_day_str = prev_day.strftime('%d-%m-%Y')

    target_folder = os.path.join("saves", today_str)
    prev_folder = os.path.join("saves", prev_day_str)

    create_folder(target_folder)

    pdf_name = find_pdf()
    rows = extract_pdf_data(pdf_name)

    prev_file = os.path.join(prev_folder, f"{prev_day_str}.xlsx")
    move_pdf(target_folder)

    # Copier le fichier stock de masse
    stock_file = "Stock de masse 2025 referance.xlsx"
    if os.path.exists(stock_file):
        shutil.copy(stock_file, os.path.join(target_folder, os.path.basename(stock_file)))
        print(f"Copie du fichier stock de masse dans {target_folder}")
    else:
        print("Fichier stock de masse introuvable, impossible de copier.")

if __name__ == "__main__":
    main()
