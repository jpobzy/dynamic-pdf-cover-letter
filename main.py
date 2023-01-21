import docx
import os
from docx.shared import Pt

def replace_word_pdf(file,  theCompany, new_pdf_name, YourName, YourEmail, YourNumber, YourWebsite):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        for run in para.runs:
            if run.text.startswith("HeaderName"):
                run.text = run.text.replace("HeaderName", YourName)
                run.font.size = Pt(25)
            elif run.text.strip().startswith("YourEmail"):
                run.text = run.text.replace("YourEmail", YourEmail)
                run.font.size = Pt(11)
            elif run.text.strip().startswith("YourNumber"):
                run.text = run.text.replace("YourNumber", YourNumber)
                run.font.size = Pt(11)
            elif run.text.strip().startswith("YourWebsite"):
                run.text = run.text.replace("YourWebsite", YourWebsite)
                run.font.size = Pt(11)
            elif run.text.strip().startswith("FooterName"):
                run.text = run.text.replace("FooterName", YourName)
                run.font.size = Pt(11)
            elif run.text != '\n':
                run.text = run.text.replace("COMPANY", theCompany)
                run.font.size = Pt(11)
    doc.save(new_pdf_name+'.docx')
    from docx2pdf import convert
    convert(new_pdf_name+'.docx',new_pdf_name+'.pdf')
    os.remove(new_pdf_name+'.docx')

file = "coverletter.docx"
company =  "THE AMONGUS COMPANY"
new_pdf_name = "new_coverletter"
your_name = "Joey Poblete"
email = "myemail@aol.com"
phone_number ="my-phone-number"
website = "my-website.com"

replace_word_pdf(file, company, new_pdf_name, your_name, email, phone_number, website)