import pandas as pd
from pathlib import Path as path
from docxtpl import DocxTemplate

# creating directory paths
base_dir = path(__file__).parent.parent
word_template_path = base_dir / "P1S945_01.docx"
excel_path = base_dir / "P1S945_01.xlsx"
reports_dir = base_dir / "reports_generator"

print(word_template_path)
print(excel_path)
print(reports_dir)

# read excel file and clean the data 
df = pd.read_excel(excel_path, sheet_name="students")
P1S945_01_dataframe = df.fillna(" ")

dict = P1S945_01_dataframe.to_dict(orient="records")
print(dict)

for record in dict:
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    reports_path = reports_dir / f"report-{record['no']}.docx"
    doc.save(reports_path)