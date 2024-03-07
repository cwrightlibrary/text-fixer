from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from os.path import dirname, join, realpath

current_dir = dirname(realpath(__file__))

original_document = Document(join(current_dir, "job-list.docx"))

for paragraph in original_document.paragraphs:
    if paragraph.text[0:11] == "Job Title: ":
        job_title = paragraph.text[11:]
    elif paragraph.text[0:17] == "Job Description: ":
        job_desc = paragraph.text[17:]
    elif paragraph.text[0:16] == "Qualifications: ":
        job_qual = paragraph.text[16:]
    elif paragraph.text[0:6] == "Type: ":
        job_type = paragraph.text[6:]
    elif paragraph.text[0:14] == "How to Apply: ":
        job_how = paragraph.text[14:]
    

print(job_how)
