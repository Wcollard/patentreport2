import os
import tkinter as tk
from tkinter import messagebox
import xlsxwriter
from datetime import datetime
import requests
from bs4 import BeautifulSoup


def get_downloads_path() -> str:
    """
    Return the path to the current user's Downloads folder in a cross-platform way.
    If Downloads does not exist, fall back to the home directory.
    """
    home = os.path.expanduser("~")
    downloads = os.path.join(home, "Downloads")
    if os.path.isdir(downloads):
        return downloads
    return home


def get_patent_details(google_url: str):
    """
    Fetch patent details from a Google Patents URL.
    Returns: (pdf_link, title, inventor, publication_date, abstract)
    """
    try:
        response = requests.get(google_url, timeout=10)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")

            # PDF link
            pdf_link_tag = soup.find("a", string="Download PDF")
            pdf_link = pdf_link_tag["href"] if pdf_link_tag else "PDF link not found"

            # Title
            title_tag = soup.find("meta", attrs={"name": "DC.title"})
            title = title_tag["content"] if title_tag else "Title not found"

            # Inventor
            inventor_tag = soup.find("meta", attrs={"name": "DC.contributor"})
            inventor = inventor_tag["content"] if inventor_tag else "Inventor not found"

            # Publication date
            pub_date_tag = soup.find("meta", attrs={"name": "DC.date"})
            publication_date = (
                pub_date_tag["content"]
                if pub_date_tag
                else "Publication date not found"
            )

            # Abstract
            abstract_tag = soup.find("meta", attrs={"name": "description"})
            abstract = (
                abstract_tag["content"].strip()
                if abstract_tag
                else "Abstract not found"
            )

            return pdf_link, title, inventor, publication_date, abstract

        # Non-200 HTTP code
        return (
            "PDF not found",
            "Title not found",
            "Inventor not found",
            "Publication date not found",
            "Abstract not found",
        )
    except Exception as e:
        # Always return 5 values to avoid unpacking errors
        return f"Error: {e}", "", "", "", ""


def export_to_excel():
    """
    Read patent numbers from the text box, fetch details, and
    export them to an Excel file in the Downloads folder.
    """
    patent_numbers = text_input.get("1.0", tk.END).strip().split("\n")
    if not patent_numbers or not any(p.strip() for p in patent_numbers):
        messagebox.showwarning(
            "No Input", "Please enter at least one patent or publication number."
        )
        return

    # Build Downloads path + timestamped filename
    downloads_path = get_downloads_path()
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(downloads_path, f"patent_export_{now}.xlsx")

    # Create workbook and worksheet
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Formats
    header_format = workbook.add_format(
        {
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "border": 1,
        }
    )

    abstract_format = workbook.add_format(
        {
            "text_wrap": True,
            "valign": "top",
            "border": 1,
        }
    )

    link_format = workbook.add_format(
        {
            "underline": True,
            "color": "blue",
            "border": 1,
        }
    )

    # Column widths
    column_widths = [20, 20, 20, 25, 15, 15, 15, 15, 70, 50]
    for col, width in enumerate(column_widths):
        worksheet.set_column(col, col, width)

    # Headers
    headers = [
        "Ref No.",
        "Inventor",
        "Title",
        "Publication Date",
        "PDF_LINK",
        "Google Link",
        "Espacenet Link",
        "USPTO Link",
        "ABSTRACT",
        "NOTES",
    ]
    worksheet.write_row(0, 0, headers, header_format)

    # Data rows
    row = 1
    for number in patent_numbers:
        cleaned_number = number.strip()
        if not cleaned_number:
            continue

        google_url = f"https://patents.google.com/patent/{cleaned_number}/en"
        espacenet_url = f"https://worldwide.espacenet.com/patent/search?q={cleaned_number}"
        uspto_number = cleaned_number.replace("US", "")
        uspto_url = f"https://ppubs.uspto.gov/pubwebapp/external.html?q={uspto_number}.pn."
        uspto_pdf_link = f"https://image-ppubs.uspto.gov/dirsearch-public/print/downloadPdf/{uspto_number}"

        pdf_link, title, inventor, publication_date, abstract = get_patent_details(
            google_url
        )

        # Write data
        worksheet.write(row, 0, cleaned_number)
        worksheet.write(row, 1, inventor)
        worksheet.write(row, 2, title, abstract_format)
        worksheet.write(row, 3, publication_date)
        worksheet.write(row, 4, pdf_link)
        worksheet.write_url(row, 5, google_url, link_format, string=cleaned_number)
        worksheet.write_url(row, 6, espacenet_url, link_format, string=cleaned_number)
        worksheet.write_url(row, 7, uspto_url, link_format, string=cleaned_number)
        worksheet.write(row, 8, abstract, abstract_format)
        worksheet.write(row, 9, "")  # NOTES column

        # Rough row height based on abstract length
        text_lines = len(abstract) // 70 + abstract.count("\n") + 1
        row_height = min(text_lines * 15, 409)  # 409 = Excel max row height
        worksheet.set_row(row, row_height)

        row += 1

    workbook.close()
    messagebox.showinfo(
        "Export Successful",
        f"Data exported to:\n{filename}",
    )


# Tkinter UI
root = tk.Tk()
root.title("Patent Exporter")

tk.Label(root, text="Enter patent/publication numbers (one per line):").pack(pady=5)
text_input = tk.Text(root, height=15, width=45)
text_input.pack(padx=10)

tk.Button(root, text="Export to Excel", command=export_to_excel).pack(pady=10)

root.mainloop()
