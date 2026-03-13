# Automation Prototype source code.
import pdfplumber
import camelot
from camelot.io import read_pdf
from docx import Document
from docx.shared import Pt
import docx
import pandas as pd
from typing import Union, Literal
from dataclasses import dataclass

# default paths we should be using for our reports, i.e. ./Reports
DEFAULT_ACTIVITY_REPORT_PATH: str = "./Reports/OrbitalFire-ActivityReportDemo.pdf" # standardize the paths
DEFAULT_FINDINGS_REPORT_PATH: str = "./Reports/Sample499/FindingsReportTest.docx" # if we're creating the report
DEFAULT_GLOSSARY_PATH: str = "./Reports/OrbitalFire-Glossary.csv"
DEFAULT_TECHNICAL_REPORT_PATH: str = "./Reports/OrbitalFire-TechnicalReportDemo.pdf"

"""
Get the report paths via user input, returns a tuple of the paths we yield.
"""
def getReports() -> tuple[str, str]:
    activity_report: str = input("Please enter the file path for Activity Report: ")
    #executive_report = input("Please enter the file path for Executive Report: ")
    #technical_report = input("Please enter the file path for Technical Report: ")
    findings_report = input("Please enter the file path for Findings Report: ")
    return (activity_report, findings_report)

"""
Find all testing from the activity report, and append it to the findings report.
"""
def automated_testing_activity(activity_report: str, findings_report: str) -> None:
    print(f"activity report path {activity_report}")
    print(f"findings report path {findings_report}")
    table_container = read_pdf(activity_report, pages="3-7", flavor="lattice")
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

# related to the excel file
def excelwork():
    print("test")
    glossary = pd.read_csv(DEFAULT_GLOSSARY_PATH)

def locate_image_technical_report(technical_report: str):
    # use pdf plumber for the other stuff, for the entries, we can get valid information
    data = pdfplumber.open(technical_report)

# type to classify our vulnerabilities
type VulnerabilityRating = Union[Literal["Informational"], Literal["Low"], Literal["Medium"], Literal["High"], Literal["Critical"]]
"""
Determines if a table entry is a vulnerability rating, targeted for certain tables
"""
def isVulnerabilityRating(val: str) -> bool:
    return val == "Informational" or val == "Low" or val == "Medium" or val == "High" or val == "Critical"

@dataclass
class VulnerabilityFrame:
    discoveredName: str
    rating: VulnerabilityRating
"""
Counts the number of severities inside of the technical report, helps get the final number we need
in the findings report
"""
def severity_counter(technical_report: str) -> list[VulnerabilityFrame]:
    table_container = read_pdf(technical_report, pages="3-6", flavor="lattice")
    frames: list[VulnerabilityFrame] = []
    # iterate over entries, find the table containing vuln classifications in frame[2], very targeted
    for entry in table_container:
        for frame in entry.df.values.tolist():
            if(len(frame) >= 3 and isVulnerabilityRating(frame[2])):
                frames.append(VulnerabilityFrame(discoveredName=frame[0], rating=frame[2]))
                print(f"vuln detected: {frame[2]}")
    # csv: risk, controls
    #      low    1
    # etc
    return frames
def put_counter_into_findings(findings_report: str, frames: list[VulnerabilityFrame]):
    # convert to csv, and inject into findings
    # findings, text box 2
    doc = Document(findings_report)
    print(len(doc.paragraphs))
    print(len(doc.inline_shapes))
    for shape in doc.inline_shapes:
        print(shape.type)



def main() -> None:
    # activity_report, findings_report = getReports()
    # automated_testing_activity(activity_report, findings_report)
    # locate_image_technical_report(DEFAULT_TECHNICAL_REPORT_PATH)
    ratings = severity_counter(DEFAULT_TECHNICAL_REPORT_PATH)
    print(f"we have {len(ratings)} vulnerabilities")
    print(ratings)
    put_counter_into_findings(DEFAULT_FINDINGS_REPORT_PATH, [])
if __name__ == "__main__":
    main()