# from pathlib import Path as path
# from docxtpl import DocxTemplate

# base_dir = path(__file__).parent.parent
# word_template_path = base_dir / "P1S945_01.docx"

# print(word_template_path) # test file path

# doc = DocxTemplate(word_template_path)
# print(doc) # test DocxTemplate

# context = {
#     "FIRSTNAME": "first_name",
#     "MIDDLENAME": "middle_name",
#     "LASTNAME": "last_name",
#     "AGE": "age",
#     "DOB": "date_of_birth"
# }
# print(context) # test data mapping


# reports_dir = base_dir / "reports"

# doc.render(context)
# doc.save(reports_dir / "generator_report.docx")