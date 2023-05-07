import datetime
import os
from pathlib import Path
import pandas as pd
import win32com.client as win32
from docxtpl import DocxTemplate
from babel.dates import format_datetime
from jinja2 import Environment, FileSystemLoader
class Generador:
    
    def __init__(self) -> None:
        self.today_date = format_datetime(datetime.datetime.today(), "d 'de' MMMM, y", locale="es")


    def documents(self, excel_file, word_template, outpur_dir):
        self.df = pd.read_excel(excel_file)
        print(self.df)
        self.word_template = word_template
        self.output_dir = Path(outpur_dir)


    def validation_info(self):
        # Load the Jinja environment and file system loader
        env = Environment(loader=FileSystemLoader(Path(self.word_template).parent))
        print(env)
        # Load the Jinja template
        template = env.get_template(os.path.basename(self.word_template))
        # Extract all the variables from the template
        self.variables = template.module.__dict__
        
        print(self.variables)


    def render_template(self, to_word:bool, to_pdf:bool):
        n = 1
        for record in self.df.to_dict(orient="records"):
            print(record)
            doc = DocxTemplate(self.word_template)
            doc.render(record)
            output_path = self.output_dir / f"carta{n}.docx"
            doc.save(output_path)
            if to_pdf:
                self.convert_to_pdf(output_path)
            if not to_word:
                os.remove(output_path)
            n += 1


    def convert_to_pdf(self, doc):
        word = win32.DispatchEx('Word.Application')
        pdf_name = str(doc).replace(".docx", r".pdf")
        word_document= word.Documents.Open(str(doc))
        word_document.SaveAs(pdf_name, FileFormat=17)
        word_document.Close()
        word.Quit()
        return None
    
# base_dir = Path(__file__).parent
    


# month_today = format_datetime(today, "MMMM", locale="es")


# excel_name = input("Escriba el nombre del excel a trabajar y presione enter: ")
# template_name = input("Escriba el nombre del word (carta) a trabajar y presione enter: ")

# template_path = base_dir / f"{template_name}.docx"
# excel_path = base_dir / f"{excel_name}.xlsx"
# output_dir = base_dir / "cartas" / f"cartas-{today_date}"

# output_dir.mkdir(exist_ok=True)


# fecha_pago = df["fecha_pago"][0]
# fecha_pago = format_datetime(fecha_pago, "d 'de' MMMM y", locale="es")
