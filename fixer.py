from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from os.path import dirname, join, realpath
from spellchecker import SpellChecker
import re

current_dir = dirname(realpath(__file__))
spell = SpellChecker()


def get_job_info(doc):
    original_doc = Document(join(current_dir, doc))

    for paragraph in original_doc.paragraphs:
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
        elif paragraph.text[0:7] == "Salary:":
            job_salary = paragraph.text[7:]
        elif paragraph.text[0:9] == "Contact: ":
            job_contact = paragraph.text[9:]

    job = {
        "title": job_title,
        "desc": job_desc,
        "qual": job_qual,
        "type": job_type,
        "how": job_how,
        "salary": job_salary,
        "contact": job_contact,
    }

    return job


original_doc = get_job_info("job-list.docx")

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

    new_element = new_element.replace("Caroliniana", "Carolina").replace("caroliniana", "carolina")
    
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
    pass

# test_file = open(join(current_dir, "test_file.txt"), "w", encoding="utf-8")
# test_file.write(testing)
# test_file.close()
