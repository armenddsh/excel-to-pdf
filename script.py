import pandas as pd
import pdfkit
import os
from glob import glob
from pathlib import Path


for filename in glob("*xls"):
    print(filename)
    current_dir = os.getcwd()

    filename = os.path.join(current_dir, filename)

    df = pd.read_excel(filename)

    filename_without_extension = os.path.join(current_dir, Path(filename).stem)
    filename_html = f"{filename_without_extension}.html"
    filename_pdf = f"{filename_without_extension}.pdf"

    df.to_html(filename_html)

    if os.name == "nt":
        config = pdfkit.configuration(wkhtmltopdf=os.path.join(current_dir, "wkhtmltopdf.exe"))
        pdfkit.from_file(filename_html, filename_pdf, configuration=config)
    else:
        pdfkit.from_file(filename_html, filename_pdf)