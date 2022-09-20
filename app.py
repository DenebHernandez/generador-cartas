import datetime
import pandas as pd
import win32com.client as win32
from pathlib import Path
from docxtpl import DocxTemplate
from babel.dates import format_datetime


today = datetime.datetime.today()
today_date = format_datetime(today, "d 'de' MMMM, y", locale="es")
month_today = format_datetime(today, "MMMM", locale="es")

base_dir = Path(__file__).parent
template_path = base_dir / "template.docx"
excel_path = base_dir / "datos.xlsx"
output_dir = base_dir / "cartas" / f"cartas-{today_date}"

output_dir.mkdir(exist_ok=True)

df = pd.read_excel(excel_path, sheet_name="Sheet1")

fecha_pago = df["fecha_pago"][0]
fecha_pago = format_datetime(fecha_pago, "d 'de' MMMM y", locale="es")

def convert_to_pdf(doc):
    word = win32.DispatchEx('Word.Application')
    pdf_name = str(doc).replace(".docx", r".pdf")
    word_document= word.Documents.Open(str(doc))
    word_document.SaveAs(pdf_name, FileFormat=17)
    word_document.Close()
    word.Quit()
    return None


for record in df.to_dict(orient="records"):
    record['fecha_actual'] = today_date
    record['mes_actual'] = month_today
    record['monto_deuda'] = '{:,.2f}'.format(record['monto_deuda'])
    record['fecha_pago'] = fecha_pago
    # print(record)
    doc = DocxTemplate(template_path)
    doc.render(record)
    output_path = output_dir / f"{record['nombre_estudiante']}-carta.docx"
    doc.save(output_path)
    convert_to_pdf(output_path)
