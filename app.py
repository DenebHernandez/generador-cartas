import datetime
import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
import win32com.client as win32

today = datetime.datetime.today().strftime("%A, %w de %B")
month_today = datetime.datetime.today().strftime("%B")

base_dir = Path(__file__).parent
template_path = base_dir / "template.docx"
excel_path = base_dir / "datos.xlsx"
output_dir = base_dir / f"cartas-{today}"

output_dir.mkdir(exist_ok=True)

df = pd.read_excel(excel_path, sheet_name="Sheet1")


def convert_to_pdf(doc):
    word = win32.DispatchEx('Word.Application')
    pdf_name = str(doc).replace(".docx", r".pdf")
    word_document= word.Documents.Open(str(doc))
    word_document.SaveAs(pdf_name, FileFormat=17)
    word_document.Close()
    return None


for record in df.to_dict(orient="records"):
    record['fecha_actual'] = today
    record['mes_actual'] = month_today
    record['monto_deuda'] = '{:,.2f}'.format(record['monto_deuda'])
    # print(record)
    doc = DocxTemplate(template_path)
    doc.render(record)
    output_path = output_dir / f"{record['nombre_estudiante']}-carta.docx"
    doc.save(output_path)
    convert_to_pdf(output_path)
