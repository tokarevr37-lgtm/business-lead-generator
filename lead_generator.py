import csv
from pathlib import Path

from openpyxl import Workbook


CSV_FILE = "business_leads.csv"
EXCEL_FILE = "business_leads.xlsx"


LEADS = [
    {
        "company_name": "Muster Clean GmbH",
        "website": "https://www.musterclean.de",
        "email": "info@musterclean.de",
        "phone": "+49 251 123456",
        "city": "Münster",
        "industry": "Cleaning Services",
    },
    {
        "company_name": "Westfalen Gebäudeservice",
        "website": "https://www.westfalen-gebaeudeservice.de",
        "email": "kontakt@westfalen-gebaeudeservice.de",
        "phone": "+49 251 987654",
        "city": "Münster",
        "industry": "Facility Services",
    },
    {
        "company_name": "CleanPro NRW",
        "website": "https://www.cleanpro-nrw.de",
        "email": "hello@cleanpro-nrw.de",
        "phone": "+49 211 445566",
        "city": "Düsseldorf",
        "industry": "Commercial Cleaning",
    },
]


def save_to_csv(data, filename):
    fieldnames = [
        "company_name",
        "website",
        "email",
        "phone",
        "city",
        "industry",
    ]

    with open(filename, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)


def save_to_excel(data, filename):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Business Leads"

    headers = ["Company Name", "Website", "Email", "Phone", "City", "Industry"]
    sheet.append(headers)

    for lead in data:
        sheet.append([
            lead["company_name"],
            lead["website"],
            lead["email"],
            lead["phone"],
            lead["city"],
            lead["industry"],
        ])

    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        sheet.column_dimensions[column_letter].width = max_length + 2

    workbook.save(filename)


def main():
    print("Preparing business lead list...")

    save_to_csv(LEADS, CSV_FILE)
    save_to_excel(LEADS, EXCEL_FILE)

    print(f"Done. Saved {len(LEADS)} leads.")
    print(f"CSV saved to: {Path(CSV_FILE).resolve()}")
    print(f"Excel saved to: {Path(EXCEL_FILE).resolve()}")


if __name__ == "__main__":
    main()
