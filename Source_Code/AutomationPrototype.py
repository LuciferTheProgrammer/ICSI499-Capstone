# Automation Prototype source code.
import pdfplumber
import camelot
from docx import Document
from docx.shared import Pt

technical_report = input("Please enter the file path for Technical Report: ")
#executive_report = input("Please enter the file path for Executive Report: ")
#activity_report = input("Please enter the file path for Activity Report: ")
findings_report = input("Please enter the file path for Findings Report: ")

def automated_testing_activity(technical_report):
    table_container = camelot.read_pdf(technical_report, pages="3-7", flavor="lattice")
    activity_log = []

    for entry in table_container:
        sample = entry.df
        for event in sample.values.tolist():
            string1 = event[0].replace("\n", "")
            string2 = event[1]
            string3 = event[2]
            cleaned = string1 + " " + string2 + " " + string3
            activity_log.append(cleaned)
    #        print(event)
    print("✅ All events copied.")
    doc_holder = Document(findings_report)
    section = "AUTOMATED TESTING ACTIVITY"
    for i, paragraph in enumerate(doc_holder.paragraphs):
        if section.lower() in paragraph.text.lower():
            for entry in reversed(activity_log):
                container = doc_holder.paragraphs[i+1].insert_paragraph_before()
                container.style = "Activity Bullet"
                execute  = container.add_run(entry)
                execute.font.size = Pt(12)
            break
    print("✅ All events pasted.")
    doc_holder.save(findings_report)
    print("✅ Document saved.")

def main():
    automated_testing_activity(technical_report)
main()