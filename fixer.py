from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from os.path import dirname, join, realpath
import re

current_dir = dirname(realpath(__file__))


def get_job_info(doc):
    original_doc = Document(join(current_dir, doc))

    for paragraph in original_doc.paragraphs:
        if "submission date" in paragraph.text.lower():
            job_date = paragraph.text[17:]
        if paragraph.text[0:17] == "Submission Date: ":
            job_date = paragraph.text[17:]
        elif paragraph.text[0:11] == "Job Title: ":
            job_title = paragraph.text[11:]
        elif paragraph.text[0:17] == "Job Description: ":
            job_desc = paragraph.text[17:]
        elif paragraph.text[0:16] == "Qualifications: ":
            job_qual = paragraph.text[16:]
        elif paragraph.text[0:6] == "Type: ":
            job_type = paragraph.text[6:]
        elif paragraph.text[0:14] == "How to Apply: ":
            job_how = paragraph.text[14:]
        elif paragraph.text[0:7] == "Salary:":
            job_salary = paragraph.text[7:]
        elif paragraph.text[0:9] == "Contact: ":
            job_contact = paragraph.text[9:]

    job = {
        "date": job_date,
        "title": job_title,
        "desc": job_desc,
        "qual": job_qual,
        "type": job_type,
        "how": job_how,
        "salary": job_salary,
        "contact": job_contact,
    }

    return job


def fix_date(element):
    new_element = element
    
    return new_element


def fix_title(element):
    new_element = element
    new_element = new_element.replace("and", "&").title()
    
    return new_element


def fix_description(element):
    new_element = element
    letter_spacing = re.findall("[a-z][A-Z]", new_element)
    colon_spacing = re.findall(":[a-zA-Z]", new_element)
    period_spacing = re.findall(".[a-zA-Z]", new_element)
    special_spacing = colon_spacing + period_spacing

    for e in special_spacing:
        if e[0] == ":":
            new_element = new_element.replace(e, e[0] + "\n" + e[1])
        elif e[0] == ".":
            new_element = new_element.replace(e, e[0] + " " + e[1])

    for e in letter_spacing:
        new_element = new_element.replace(e, e[0] + ".\n" + e[1])

    if new_element[len(new_element) - 1:] != ".":
        new_element += "."
    
    return new_element


def fix_qualifications(element):
    new_element = element
    letter_spacing = re.findall("[a-z][A-Z]", new_element)
    colon_spacing = re.findall(":[a-zA-Z]", new_element)
    period_spacing = re.findall(".[a-zA-Z]", new_element)
    special_spacing = colon_spacing + period_spacing

    for e in special_spacing:
        if e[0] == ":":
            new_element = new_element.replace(e, e[0] + "\n" + e[1])
        elif e[0] == ".":
            new_element = new_element.replace(e, e[0] + " " + e[1])

    for e in letter_spacing:
        new_element = new_element.replace(e, e[0] + ".\n" + e[1])

    if new_element[len(new_element) - 1:] != ".":
        new_element += "."
    
    return new_element


def fix_type(element):
    new_element = element
    if "part" in new_element or "Part" in new_element:
        new_element = "Part-Time"
    elif "full" in new_element or "Full" in new_element:
        new_element = "Full-Time"
    else:
        new_element = new_element.title()
    
    return new_element


def fix_salary(element):
    new_element = element
    if new_element == "":
        new_element = ""
    
    return new_element


def fix_contact(element):
    new_element = element
    letter_spacing = re.findall("[a-z][A-Z]", new_element)
    
    for e in letter_spacing:
        if e[0] != "h" and e[1] != "D":
            new_element = new_element.replace(e, e[0] + " " + e[1])
    
    phone = re.compile(r'\b(?:\d{3}[-.\s]?)?\d{3}[-.\s]?\d{4}\b').search(new_element)
    
    if phone:
        phone = phone.group()
        new_element = new_element.replace(" " + phone, ", " + phone)
    else:
        phone = ""
    
    return new_element


def load_fix_document(loc):
    original_doc = get_job_info(loc)

    new_doc = {}
    new_doc["Submission Date: "] = fix_date(original_doc["date"])
    new_doc["Job Title: "] = fix_title(original_doc["title"])
    new_doc["Job Description: "] = fix_description(original_doc["desc"])
    new_doc["Qualifications: "] = fix_qualifications(original_doc["qual"])
    new_doc["Type: "] = fix_type(original_doc["type"])
    new_doc["How to Apply: "] = original_doc["how"]
    new_doc["Salary: "] = fix_salary(original_doc["salary"])
    new_doc["Contact: "] = fix_contact(original_doc["contact"])
    
    return new_doc


def create_new_document(new_doc):
    new_document = Document()
    style = new_document.styles["Normal"]
    font = style.font
    font.name = "Arial"
    hstyle = new_document.styles["Heading 2"]
    hstyle.font.name = "Arial"
    hstyle.font.size = Pt(16)
    
    h = new_document.add_heading("", 2)
    h.add_run(new_doc["Job Title: "] + ", " + new_doc["Submission Date: "]).bold = False
    s = new_document.add_paragraph()
    
    for k, v in new_doc.items():
        p = new_document.add_paragraph()
        p.style = new_document.styles["Normal"]
        p.add_run(k).bold = True
        p.add_run(v)
        p.add_run()
    
    return new_document


fix_document = load_fix_document("job-list.docx")
new_document = create_new_document(fix_document)

new_document.save(join(current_dir, "test.docx"))
