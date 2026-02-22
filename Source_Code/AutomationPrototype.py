# Automation Prototype source code.
import pdfplumber
import camelot
from docx import Document
from docx.shared import Pt

DEFAULT_ACTIVITY_REPORT_PATH: str = "./Reports/ActivityReport.pdf" # standardize the paths
DEFAULT_FINDINGS_REPORT_PATH: str = "./Reports/FindingsReport.pdf" # if we're creating the report

def getReports() -> tuple[str, str]:
    activity_report: str = input("Please enter the file path for Activity Report: ")
    #executive_report = input("Please enter the file path for Executive Report: ")
    #technical_report = input("Please enter the file path for Technical Report: ")
    findings_report = input("Please enter the file path for Findings Report: ")
    return (activity_report, findings_report)

def automated_testing_activity(activity_report: str, findings_report: str) -> None:
    table_container = camelot.read_pdf(activity_report, pages="3-7", flavor="lattice")
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
                execute.font.size = Pt(8.5)
            break
    print("✅ All events pasted.")
    doc_holder.save(findings_report)
    print("✅ Document saved.")

def main() -> None:
    activity_report, findings_report = getReports()
    automated_testing_activity(activity_report, findings_report)
if __name__ == "__main__":
    main()