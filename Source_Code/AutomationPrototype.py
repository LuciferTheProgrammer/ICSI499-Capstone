# Automation Prototype source code.
import base64
import pdfplumber
import camelot
from camelot.io import read_pdf
from docx import Document
from docx.shared import Pt, Inches
import pandas as pd
from typing import Union, Literal
from dataclasses import dataclass
import os
import json
from openai import OpenAI
import pymupdf


# default paths we should be using for our reports, i.e. ./Reports
DEFAULT_ACTIVITY_REPORT_PATH: str = "./Reports/OrbitalFire-ActivityReportDemo.pdf" # standardize the paths
DEFAULT_FINDINGS_REPORT_PATH: str = "./Reports/Sample499/FindingsReportTest.docx" # if we're creating the report
DEFAULT_GLOSSARY_PATH: str = "./Reports/OrbitalFire-Glossary.csv"
DEFAULT_TECHNICAL_REPORT_PATH: str = "./Reports/OrbitalFire-TechnicalReportDemo.pdf"
DEFAULT_EXECUTIVE_REPORT_PATH: str = "./Reports/OrbitalFire-ExecutiveReportDemo.pdf"
DEFAULT_RECOMMENDATIONS_PATH: str = "./Reports/FindingsDetailsAndRecommendations.xlsx"

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

# Populates the Assessment Results Summary section in the Findings Report.
def assessment_results(executive_report: str, findings_report: str) -> None:
    client_call = OpenAI(api_key = os.environ.get("OPENAI_API_KEY"))
    pdf = pymupdf.open(executive_report)
    page_n = pdf[5]
    pixels = page_n.get_pixmap(matrix=pymupdf.Matrix(2,2))
    image_size = pixels.tobytes("png")
    image_64 = base64.b64encode(image_size).decode("utf-8")
    pdf.close()
    task = [{"role": "user", "content": [{"type": "input_text", "text": """
                            Extract the Engagement Results Summary table from this report page. Return JSON only in this exact format:
                            {
                                "items": [
                                    {
                                    "category": "...", 
                                    "summary": "..."
                                    }
                                ]
                            } 
                            Rules:
                            - Use only what is visible in the image.
                            - Extract all visible rollup findings from the Engagement Results Summary table.
                            - Pair each category with the correct summary.
                            - Merge/combine wrapped text into one summary per category.
                            - Ignore page titles and unrelated sections.
                            - Remove header rows like Category and Summary.
                            - Keep maximum 5 items.
                            - Do not invent or add new content.
                            """},
                                        {
                                            "type": "input_image",
                                            "image_url": f"data:image/png;base64,{image_64}"
                                        }
                                        ]}]
    responses = client_call.responses.create(model="gpt-5-mini", input = task)
    try :
        print(responses.output_text)
        extracted_info = json.loads(responses.output_text.strip())
        print(json.dumps(extracted_info, indent = 2))
        item_collection = extracted_info.get("items", [])
    except Exception as e:
        print("OpenAI API call returned an error")
        return
    view = set()
    clean = []
    for i in item_collection:
        final_cat = " ".join(str(i.get("category", "")).split()).strip()
        final_sum = " ".join(str(i.get("summary", "")).split()).strip()
        if not final_cat or not final_sum:
            continue
        k = (final_cat.lower(), final_sum.lower())
        if k in view:
            continue
        view.add(k)
        clean.append((final_cat, final_sum))
        computed_len = len(clean)
        if computed_len == 5:
            break
    if not clean:
        print("No valid entries found from OpenAI")
        return
    doc = Document(findings_report)
    pos = None
    s_header = "ASSESSMENT RESULTS SUMMARY"
    for j, paragraph in enumerate(doc.paragraphs):
        cleaned_p = paragraph.text.upper()
        if s_header in cleaned_p:
            pos = j
            break
    if pos is None:
        print(f"{s_header} was not found in {findings_report}")
        return
    for k in range(pos, min(pos + 8, len(doc.paragraphs))):
        print(k, repr(doc.paragraphs[k].text))
    add_pos = doc.paragraphs[pos + 2]
    for final_cat, final_sum in clean:
        cat_cont1 = add_pos.insert_paragraph_before()
        cat_cont2 = cat_cont1.add_run(final_cat.upper())
        cat_cont2.bold = True
        cat_cont2.font.name = "Corbel"
        cat_cont2.font.size = Pt(13)
        sum_cont1 = add_pos.insert_paragraph_before()
        sum_cont2 = sum_cont1.add_run(final_sum)
        sum_cont2.font.name = "Corbel"
        sum_cont2.font.size = Pt(12)
    doc.save(findings_report)
    print("✅ Assessment Results Summary saved.")

# Populates the Recommendations section in the Findings Report.
def recommendations(reco_findings: str, findings_report: str) -> None:
    pointer = pd.read_excel(reco_findings, sheet_name="External")
    collections = []
    for k, entry in pointer.iterrows():
        recommend_data = str(entry["Recommendation"]).strip()
        risk_level  =str(entry["Risk Level"]).strip().title()
        if recommend_data.lower() == "nan" or not recommend_data:
            continue
        if risk_level.lower() == "nan" or not risk_level:
            continue
        findings_num  = k + 1
        collections.append((findings_num, risk_level, recommend_data))
    if not collections:
        print("No Recommendations found column")
        return
    doc = Document(findings_report)
    pos = None
    section_header = "DISCRETIONARY REMEDIATION"
    for i, paragraph in enumerate(doc.paragraphs):
        if section_header in paragraph.text.upper():
            pos = i
            break
    if pos is None:
        print(f"{section_header} was not found in {findings_report}")
        return
    increment = doc.paragraphs[pos + 1]
    for finding_num, risk_level, recommend_data in collections:
        body_cont1 = increment.insert_paragraph_before()
        body_cont2 = body_cont1.add_run(f"{finding_num}. {recommend_data}")
        body_cont2.font.name = "Corbel"
        body_cont2.font.size = Pt(12)
        body_cont1.paragraph_format.left_indent = Inches(0.4)
        body_cont1.paragraph_format.first_line_indent = Inches(-0.2)
        body_cont1.paragraph_format.space_before = Pt(0)
        body_cont1.paragraph_format.space_after = Pt(0)
        rec_cont1 = increment.insert_paragraph_before()
        rec_cont2 = rec_cont1.add_run(f"Refer to {risk_level} Finding {finding_num}")
        rec_cont2.font.name = "Corbel"
        rec_cont2.font.size = Pt(12)
        rec_cont2.underline = True
        rec_cont1.paragraph_format.left_indent = Inches(0.4)
        rec_cont1.paragraph_format.first_line_indent = Inches(0)
        rec_cont1.paragraph_format.space_before = Pt(0)
        rec_cont1.paragraph_format.space_after = Pt(6)
    doc.save(findings_report)
    print("✅ Recommendations was saved.")

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
    return frames

def main() -> None:
    # activity_report, findings_report = getReports()
    automated_testing_activity(DEFAULT_ACTIVITY_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    assessment_results(DEFAULT_EXECUTIVE_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    recommendations(DEFAULT_RECOMMENDATIONS_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    # locate_image_technical_report(DEFAULT_TECHNICAL_REPORT_PATH)
    #ratings = severity_counter(DEFAULT_TECHNICAL_REPORT_PATH)
    #print(f"we have {len(ratings)} vulnerabilities")
    #print(ratings)
if __name__ == "__main__":
    main()