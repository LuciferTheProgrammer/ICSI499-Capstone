# Automation Prototype source code.
import base64
import re
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
    executive_report = input("Please enter the file path for Executive Report: ")
    technical_report = input("Please enter the file path for Technical Report: ")
    findings_report = input("Please enter the file path for Findings Report: ")
    return (activity_report, findings_report)

"""
Find all testing from the activity report, and append it to the findings report.
"""
def automated_testing_activity(activity_report: str, findings_report: str, page_range: str) -> None:
    table_container = read_pdf(activity_report, pages=page_range, flavor="lattice")
    activity_log = []
    for entry in table_container:
        sample = entry.df
        for event in sample.values.tolist():
            string1 = event[0].replace("\n", "")
            string2 = event[1]
            string3 = event[2]
            cleaned = string1 + " " + string2 + " " + string3
            activity_log.append(cleaned)
    print("✅ All events copied.")
    doc_holder = Document(findings_report)
    section = "AUTOMATED TESTING ACTIVITY"
    for i, paragraph in enumerate(doc_holder.paragraphs):
        if section.lower() in paragraph.text.lower():
            for entry in reversed(activity_log):
                container = doc_holder.paragraphs[i + 1].insert_paragraph_before()
                container.style = "Activity Bullet"
                execute  = container.add_run(entry)
                execute.font.name = "Corbel"
                execute.font.size = Pt(8.5)
            break
    print("✅ All events pasted.")
    doc_holder.save(findings_report)
    print("✅ Document saved.")

# Populates the Assessment Results Summary section in the Findings Report.
def assessment_results(executive_report: str, findings_report: str, page: str) -> None:
    client_call = OpenAI(api_key = os.environ.get("OPENAI_API_KEY"))
    pdf = pymupdf.open(executive_report)
    page_int = int(page)
    indexed = page_int - 1
    page_n = pdf[indexed]
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
                            - Keep the maximum 5 items.
                            - Do not invent or add new content.
                            """},
                                        {
                                            "type": "input_image",
                                            "image_url": f"data:image/png;base64,{image_64}"
                                        }
                                        ]}]
    responses = client_call.responses.create(model="gpt-5-mini", input = task)
    try :
        extracted_info = json.loads(responses.output_text.strip())
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


def process_titles_rec(x ,y):
    x = y.lower()
    x = x.replace("\n", " ")
    x = " ".join(x.split())
    x = x.replace("–", "-").replace("—", "-")
    x = x.strip()
    return x

# Populates the Recommendations section in the Findings Report.
def recommendations(recommendation_csv: str, technical_report: str, findings_report: str, environment: str) -> None:
    #environment = input("Please select the environment type: 'Internal' or 'External': ")
    data_container = pd.read_excel(recommendation_csv, sheet_name=environment)
    recommendation_map = {}
    for _, i in data_container.iterrows():
        findings_title = i["Finding Title"]
        recommendation_data = i["Recommendation"]
        if findings_title is None:
            continue
        if recommendation_data is None:
            continue
        findings_title = str(findings_title).strip()
        recommendation_data = str(recommendation_data).strip()
        if findings_title == "":
            continue
        if recommendation_data == "" or recommendation_data == "nan":
            continue
        cleaned_titles = ""
        cleaned_titles = process_titles_rec(cleaned_titles, findings_title)
        recommendation_map[cleaned_titles] = recommendation_data
    pages_content = []
    technical_finds = []
    with pdfplumber.open(technical_report) as pdf:
        for page in pdf.pages:
            page_content = page.extract_text()
            if not page_content:
                continue
            pages_content.append(page_content)
            for entry in page_content.splitlines():
                normalized = entry.strip()
                if normalized == "":
                    continue
                matching = re.match(r"^(CRITICAL|HIGH|MEDIUM|LOW|INFORMATIONAL)\s+(.+)$", normalized, re.IGNORECASE)
                if matching:
                    status = matching.group(1).title()
                    finding_title = matching.group(2).strip()
                    if finding_title == "":
                        continue
                    cleaned_title = ""
                    cleaned_title = process_titles_rec(cleaned_title, finding_title)
                    technical_finds.append({"severity": status, "finding_title": finding_title, "cleaned_title": cleaned_title})
    recommendations = []
    viewed = set()
    status_counter = {"Critical": 0, "High": 0, "Medium": 0, "Low": 0, "Informational": 0}
    for finding in technical_finds:
        processed_title = finding["cleaned_title"]
        if processed_title in viewed:
            continue
        viewed.add(processed_title)
        if processed_title not in recommendation_map:
            print(f"No recommendation found for {finding['finding_title']}")
            continue
        status_counter[finding["severity"]] = status_counter[finding["severity"]] + 1
        match = recommendation_map[processed_title]
        recommendations.append({"severity": finding["severity"], "tracker_finding": status_counter[finding["severity"]], "finding_title": finding["finding_title"], "recommendation_data": match})
    doc = Document(findings_report)
    position = None
    landing_section = "DISCRETIONARY REMEDIATION"
    for index, p in enumerate(doc.paragraphs):
        if landing_section in p.text.upper():
            position = index
            break
    if position is None:
        print(f"Recommendation section not found in {findings_report}")
        return
    len_paragraphs = len(doc.paragraphs)
    increment_position = position + 1
    if len_paragraphs > increment_position:
        new_position = doc.paragraphs[increment_position]
    else:
        new_position = doc.paragraphs[position]

    for i in recommendations:
        status_holder = i["severity"]
        tracker_holder = i["tracker_finding"]
        data_holder = i["recommendation_data"]
        left_component = data_holder
        right_component = ""
        if " - " in data_holder:
            left_component, right_component = data_holder.split(" - ", 1)
        elif " – " in data_holder:
            left_component, right_component = data_holder.split(" – ", 1)
        elif "-" in data_holder:
            left_component, right_component = data_holder.split("-", 1)
        left_component = left_component.strip()
        right_component = right_component.strip()
        body_cont1 = new_position.insert_paragraph_before()
        body_cont1.style = "List Number 2"
        body_cont2 = body_cont1.add_run(left_component)
        body_cont2.bold = True
        body_cont2.font.name = "Corbel"
        body_cont2.font.size = Pt(12)
        if right_component:
            normalized_def = body_cont1.add_run(" - " + right_component)
            normalized_def.font.name = "Corbel"
            normalized_def.font.size = Pt(12)
        rec_cont1 = new_position.insert_paragraph_before()
        rec_cont2 = rec_cont1.add_run(f"Refer to {status_holder} Finding {tracker_holder}")
        rec_cont2.font.name = "Corbel"
        rec_cont2.font.size = Pt(12)
        rec_cont2.underline = True
        rec_cont1.paragraph_format.left_indent = Inches(0.5)
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
    #recommendations(DEFAULT_RECOMMENDATIONS_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    recommendations(DEFAULT_RECOMMENDATIONS_PATH, DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    # locate_image_technical_report(DEFAULT_TECHNICAL_REPORT_PATH)
    #ratings = severity_counter(DEFAULT_TECHNICAL_REPORT_PATH)
    #print(f"we have {len(ratings)} vulnerabilities")
    #print(ratings)
if __name__ == "__main__":
    main()