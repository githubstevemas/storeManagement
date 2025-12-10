from datetime import datetime, timedelta

from django.shortcuts import render
from django.http import HttpResponse
from django.conf import settings
import os

from .utils.xlsx_processing import save_excel
from .utils.pdf_processing import extract_pdf_data
from .utils.folder import get_dates


def previous_workday(start):
    current = start - timedelta(days=1)
    while current.weekday() >= 5:  # 5=Samedi, 6=Dimanche
        current -= timedelta(days=1)
    return current


def next_workdays(start, count):
    days = []
    current = start
    while len(days) < count:
        current += timedelta(days=1)
        if current.weekday() < 5:
            days.append(current)
    return days


def process_upload(request):
    today = datetime.today()

    workdays_j = next_workdays(today, 3)
    pdf_j_label = f"{workdays_j[0].strftime('%d/%m')} au {workdays_j[-1].strftime('%d/%m')}"

    prev_workday = previous_workday(today)
    workdays_j1 = next_workdays(prev_workday, 3)
    pdf_j1_label = f"{workdays_j1[0].strftime('%d/%m')} au {workdays_j1[-1].strftime('%d/%m')}"

    context = {
        "pdf_j_label": pdf_j_label,
        "pdf_j1_label": pdf_j1_label
    }

    if request.method == "POST":
        print("FILES:", request.FILES)

        try:
            pdf_j = request.FILES.getlist("pdf_j")[0]
            pdf_j1 = request.FILES.getlist("pdf_j1")[0]
            xlsx_file = request.FILES.getlist("xlsx_file")[0]
        except IndexError:
            context["error"] = "Veuillez uploader le PDF du jour, le PDF J-1 et le fichier stock de masse."
            return render(request, "upload/upload.html", context)

        os.makedirs(settings.MEDIA_ROOT, exist_ok=True)

        pdf_j_path = os.path.join(settings.MEDIA_ROOT, pdf_j.name)
        pdf_j1_path = os.path.join(settings.MEDIA_ROOT, pdf_j1.name)
        xlsx_path = os.path.join(settings.MEDIA_ROOT, xlsx_file.name)

        for path, file in [(pdf_j_path, pdf_j), (pdf_j1_path, pdf_j1), (xlsx_path, xlsx_file)]:
            with open(path, "wb+") as f:
                for chunk in file.chunks():
                    f.write(chunk)

        try:
            rows_j = extract_pdf_data(pdf_j_path)   # [[date, ref, qty, cmd, adresses], ...]
            rows_j1 = extract_pdf_data(pdf_j1_path)

            print("rows_j:", rows_j[:5])
            print("rows_j1:", rows_j1[:5])

            # Détection des nouveaux produits
            pdf_j1_set = {(row["code"], row["qty"]) for row in rows_j1}
            new_rows = [r for r in rows_j if
                        (r["code"], r["qty"]) not in pdf_j1_set]

            today, _ = get_dates()
            today_str = today.strftime("%d-%m-%Y")

            excel_stream = save_excel(new_rows, stock_file=xlsx_path)

            response = HttpResponse(
                excel_stream,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = f'attachment; filename="{today_str}.xlsx"'
            return response

        except Exception as e:
            context["error"] = f"Erreur : {e}"
            print("[ERROR]", e)

    return render(request, "upload/upload.html", context)
