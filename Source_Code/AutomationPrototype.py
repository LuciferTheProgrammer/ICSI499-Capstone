# Automation Prototype source code.
import pdfplumber
import camelot
import re
from camelot.io import read_pdf
from docx import Document
from docx.shared import Pt
import pandas as pd
from typing import Union, Literal
from dataclasses import dataclass
from pathlib import Path

# default paths we should be using for our reports, i.e. ./Reports
DEFAULT_ACTIVITY_REPORT_PATH = "./Source_Code/Reports/OrbitalFire-ActivityReportDemo.pdf"
DEFAULT_FINDINGS_REPORT_PATH = "./Source_Code/Reports/Sample499/FindingsReportTest.docx"
DEFAULT_GLOSSARY_PATH = "./Source_Code/Reports/OrbitalFire-Glossary.csv"
DEFAULT_TECHNICAL_REPORT_PATH = "./Source_Code/Reports/OrbitalFire-TechnicalReportDemo.pdf"
DEFAULT_EXECUTIVE_REPORT_PATH = "./Source_Code/Reports/OrbitalFire-ExecutiveReportDemo.pdf"
DEFAULT_FINDINGS_DETAILS_PATH = "./Source_Code/Reports/FindingsDetailsAndRecommendations.xlsx"

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

def assessment_results(executive_report: str, findings_report: str) -> None:
    collection = []
    tabular = camelot.read_pdf(executive_report, pages="6", flavor="lattice", process_background=True, line_scale=40)
    if tabular.n == 0:
        print("No target tables for Engagement Results Summary found - page 6")
        return
    container = None
    for i in tabular:
        data_container = i.df
        flatten = " ".join(data_container.astype(str).values.flatten()).lower()
        if "category" in flatten and "summary" in flatten:
            container = data_container
            break
    if container is None:
        print("Found target table, but not Engagement Results Summary {Category + Summary} - page 6")
        return
    header_index = None
    length_target = len(container)
    for row in range(length_target):
        row_entry = " ".join([str(x).strip().lower() for x in container.iloc[row].tolist()])
        if "category" in row_entry and "summary" in row_entry:
            header_index = row
            break
    if header_index is None:
        print("Could not located header index inside detected target table")
        return
    raw_data = container.iloc[header_index + 1:].values.tolist()
    category_cur = ""
    summary_cur = ""
    for i in raw_data:
        length = len(i)
        category = (i[0] if length > 0 else "")
        summary_1 = (i[1] if length > 1 else "")
        summary_2 = (i[2] if length > 2 else "")
        parts = []
        if length > 1 and str(summary_1).strip():
            parts.append(str(summary_1).strip())
        if length > 2 and str(summary_2).strip():
            parts.append(str(summary_2).strip())
        cleaned_sum = " ".join(parts).strip()
        category_final = str(category).replace("\n", " ").strip()
        summary_final = str(cleaned_sum).replace("\n", " ").strip()
        if category_final:
            if category_cur and summary_cur:
                new_category = category_cur.strip()
                new_summary = summary_cur.strip()
                collection.append((new_category, new_summary))
                len_collection = len(collection)
                if len_collection == 5:
                    break
            category_cur = category_final
            if summary_final:
                summary_cur = summary_final
            else:
                summary_cur = ""
            continue
        if category_cur and summary_final:
            summary_cur = (summary_cur + " " + summary_final).strip() if summary_cur else summary_final
    len_collection = len(collection)
    if len_collection < 5 and category_cur and summary_cur:
        clean_cat = category_cur.strip()
        clean_sum = summary_cur.strip()
        collection.append((clean_cat, clean_sum))
    doc = Document(findings_report)
    mark = "ASSESSMENT RESULTS SUMMARY"
    insert_data = None
    for i, paragraph in enumerate(doc.paragraphs):
        if mark.lower() in paragraph.text.lower():
            insert_data = i
            break
    if insert_data is None:
        print(f"The {mark} section not found")
        return
    position = doc.paragraphs[insert_data + 2]
    for(category, summary) in reversed(collection):
        holder_cat = position.insert_paragraph_before()
        holder_cat.style = "Normal"
        container_cat = holder_cat.add_run(category)
        container_cat.bold = True
        container_cat.font.size = Pt(13)
        container_cat.font.name = "Corbel"
        holder_sum = position.insert_paragraph_before()
        container_sum = holder_sum.add_run(summary)
        container_sum.font.size = Pt(12)
        container_sum.font.name = "Corbel"
    doc.save(findings_report)
    print("✅ Assessment Results Summary saved.")

# related to the excel file
def excelwork():
    print("test")
    glossary = pd.read_csv(DEFAULT_GLOSSARY_PATH)

def locate_image_technical_report(technical_report: str):
    # use pdf plumber for the other stuff, for the entries, we can get valid information
    data = pdfplumber.open(technical_report)

# type to classify our vulnerabilities
VulnerabilityRating = Union[Literal["Informational"], Literal["Low"], Literal["Medium"], Literal["High"], Literal["Critical"]]
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

def normalize_title(title: str) -> str:
    return " ".join(str(title).strip().lower().split())

def load_findings_lookup(details_excel: str, sheet_name: str = "External") -> dict:
    df = pd.read_excel(details_excel, sheet_name=sheet_name)

    # clean column names
    df.columns = [str(col).strip() for col in df.columns]

    lookup = {}

    for _, row in df.iterrows():
        finding_title = str(row.get("Finding Title", "")).strip()
        if not finding_title:
            continue

        key = normalize_title(finding_title)

        lookup[key] = {
            "title": finding_title,
            "description": str(row.get("Finding Description", "")).strip(),
            "risk": str(row.get("Risk Level", "")).strip(),
            "recommendation": str(row.get("Recommendation", "")).strip(),
        }

    return lookup

def finding_details(technical_report: str, findings_report: str, details_excel: str, sheet_name: str = "External") -> None:
    vulnerabilities = severity_counter(technical_report)
    lookup = load_findings_lookup(details_excel, sheet_name=sheet_name)

    doc = Document(findings_report)
    section = "FINDINGS DETAILS"

    insert_index = None
    for i, paragraph in enumerate(doc.paragraphs):
        if section.lower() in paragraph.text.lower():
            insert_index = i
            break

    if insert_index is None:
        print(f"The {section} section was not found.")
        return

    insert_position = doc.paragraphs[insert_index + 1]

    for vuln in reversed(vulnerabilities):
        key = normalize_title(vuln.discoveredName)

        if key in lookup:
            detail = lookup[key]

            # Title
            title_para = insert_position.insert_paragraph_before()
            title_run = title_para.add_run(detail["title"])
            title_run.bold = True
            title_run.font.size = Pt(13)

            # Risk Level
            risk_para = insert_position.insert_paragraph_before()
            risk_run = risk_para.add_run(f"Risk Level: {detail['risk']}")
            risk_run.bold = True
            risk_run.font.size = Pt(11)

            # Description
            desc_para = insert_position.insert_paragraph_before()
            desc_para.add_run("Finding Description: ").bold = True
            desc_para.add_run(detail["description"])

            # Recommendation
            rec_para = insert_position.insert_paragraph_before()
            rec_para.add_run("Recommendation: ").bold = True
            rec_para.add_run(detail["recommendation"])

            # spacing
            insert_position.insert_paragraph_before().add_run("")

        else:
            print(f"⚠ No match found for: {vuln.discoveredName}")

        doc.save(findings_report)
    print("✅ Findings Details saved.")

def customer_name() -> str:
    name = input("Customer Name: ").strip()
    return name


def set_customer_name(findings_report_path: str) -> None:
    name = customer_name()
    if name == "":
        print("❌ No name was provided.")
        return

    doc = Document(findings_report_path)
    target = "[CUSTOMER NAME]"
    replaced = False

    for p in doc.paragraphs:
        if target in p.text:
            updated_text = p.text.replace(target, name)
            p.clear()
            run = p.add_run(updated_text)
            run.font.name = "Corbel"
            run.font.size = Pt(16)
            replaced = True

    doc.save(findings_report_path)

    if replaced:
        print("✅ Customer Name was successfully updated")
    else:
        print("❌ No [CUSTOMER NAME] placeholder was found")


def IPAddress(technical: str, findings_report_path: str) -> None:
    collector = []
    in_scope = False
    in_ip_section = False

    scope_header = "Engagement Scope of Work"
    ip_header = "IP ADDRESSES & RANGES"

    with pdfplumber.open(technical) as pdf:
        for page in pdf.pages:
            content = page.extract_text()
            if not content:
                continue

            for line in content.splitlines():
                cleaned = " ".join(line.split()).strip()

                if cleaned == "":
                    continue

                if scope_header in cleaned:
                    in_scope = True
                    continue

                if not in_scope:
                    continue

                if ip_header in cleaned:
                    in_ip_section = True
                    continue

                if in_ip_section and (
                        "Agent Information" in cleaned or
                        "Task Performed" in cleaned or
                        "Rules of Engagement" in cleaned
                ):
                    in_ip_section = False
                    break

                if in_ip_section:
                    found_ips = re.findall(
                        r"\b(?:\d{1,3}\.){3}\d{1,3}(?:/\d{1,2})?\b",
                        cleaned
                    )
                    for ip in found_ips:
                        if ip not in collector:
                            collector.append(ip)

    if not collector:
        print("❌ Could not find IP Addresses in Technical Report")
        return

    doc = Document(findings_report_path)
    target = "The following IP Addresses and hosts were evaluated during automated testing:"
    position = None

    for i, p in enumerate(doc.paragraphs):
        if target.lower() in p.text.lower():
            position = i
            break

    if position is None:
        print("❌ Could not find IP Address section in Findings Report")
        return

    insert_after = doc.paragraphs[position + 1]

    for ip in reversed(collector):
        para = insert_after.insert_paragraph_before()
        para.style = "IP"
        run = para.add_run(ip)
        run.font.name = "Corbel"
        run.font.size = Pt(12)

    doc.save(findings_report_path)
    print("✅ IP Address section populated and saved")


def Logistics(technical_path: str, findings_report: str) -> None:
    mitre_rows = []
    escalation = None

    with pdfplumber.open(technical_path) as pdf:
        in_mitre = False
        in_contact = False

        for page in pdf.pages:
            content = page.extract_text()
            if not content:
                continue

            for line in content.splitlines():
                cleaned = " ".join(line.split()).strip()

                if cleaned == "":
                    continue

                if "Primary Point of Contact" in cleaned:
                    in_contact = True
                    continue

                if in_contact and cleaned.startswith("Name:"):
                    escalation = cleaned.replace("Name:", "").strip()
                    in_contact = False

                if "MITRE ATT&CK Mappings" in cleaned:
                    in_mitre = True
                    continue

                if in_mitre and "Reputational Threat Findings" in cleaned:
                    in_mitre = False

                if in_mitre:
                    if cleaned.startswith("Time Name Tactic"):
                        continue

                    if re.search(
                            r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}',
                            cleaned
                    ):
                        mitre_rows.append(cleaned)

    if not mitre_rows:
        print("❌ Could not find dates in MITRE section")
        return

    if not escalation:
        print("❌ Could not find escalation contact")
        return

    first = mitre_rows[0]
    last = mitre_rows[-1]

    start_match = re.search(
        r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}',
        first
    )
    end_match = re.search(
        r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}',
        last
    )

    if not start_match or not end_match:
        print("❌ Could not extract start/end dates")
        return

    start_date = start_match.group(0)
    end_date = end_match.group(0)

    doc = Document(findings_report)
    marker = "The Penetration Testing was conducted according to the following:"
    index = None

    for i, p in enumerate(doc.paragraphs):
        if marker in p.text:
            index = i
            break

    if index is None:
        print("❌ Could not find Logistics section in Findings Report")
        return

    rows = [
        (doc.paragraphs[index + 1], "Start Date: ", start_date),
        (doc.paragraphs[index + 2], "End Date: ", end_date),
        (doc.paragraphs[index + 3], "Escalation Contact: ", escalation),
    ]

    for para, label, value in rows:
        para.text = label
        run = para.add_run(value)
        run.font.name = "Corbel"
        run.font.size = Pt(12)

    doc.save(findings_report)
    print("✅ Logistics section populated and saved")


def main() -> None:
    # activity_report, findings_report = getReports()
    set_customer_name(DEFAULT_FINDINGS_REPORT_PATH)
    automated_testing_activity(DEFAULT_ACTIVITY_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    assessment_results(DEFAULT_EXECUTIVE_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    IPAddress(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    Logistics(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    finding_details(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH, DEFAULT_FINDINGS_DETAILS_PATH, "External")
    # locate_image_technical_report(DEFAULT_TECHNICAL_REPORT_PATH)
    ratings = severity_counter(DEFAULT_TECHNICAL_REPORT_PATH)
    print(f"we have {len(ratings)} vulnerabilities")
    print(ratings)
if __name__ == "__main__":
    main()