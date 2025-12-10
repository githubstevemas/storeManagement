import os

import pdfplumber
import re


def find_pdf(folder='.'):
    pdf_name = next((f for f in os.listdir(folder) if f.lower().endswith('.pdf')), None)
    if not pdf_name:
        raise FileNotFoundError("Aucun PDF trouvé dans le dossier courant.")
    return pdf_name


def extract_pdf_data(pdf_name):
    pattern_ref = re.compile(r"^(\d+)\s+([A-Z0-9]+)\s+(.+)$")
    pattern_cmd = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(\d+).+?(\d{6})")

    rows = []
    current_ref = None

    with pdfplumber.open(pdf_name) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue

                m_ref = pattern_ref.match(line)
                if m_ref:
                    current_ref = m_ref.group(2)
                    continue

                m_cmd = pattern_cmd.match(line)
                if m_cmd and current_ref:
                    date_str = m_cmd.group(1)
                    qty = int(m_cmd.group(2))
                    cmd = m_cmd.group(3)

                    rows.append({
                        "date": date_str,
                        "code": current_ref,
                        "qty": qty,
                        "cmd": cmd
                    })

    return rows

