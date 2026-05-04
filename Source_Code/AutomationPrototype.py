# Automation Prototype source code.
import base64
import win32com.client
import re
import zipfile
import shutil
import tempfile
from pathlib import Path
from collections import Counter
from operator import truediv
import openpyxl
from lxml import etree
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
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from PIL import Image
from docx.shared import RGBColor
import matplotlib.pyplot as plt
from openpyxl.utils import get_column_letter


# from packaging.utils import NormalizedName

# default paths we should be using for our reports, i.e. ./Reports
DEFAULT_ACTIVITY_REPORT_PATH: str = "./Reports/OrbitalFire-ActivityReportDemo.pdf" # standardize the paths
DEFAULT_FINDINGS_REPORT_PATH: str = "./Reports/Sample499/FindingsReportTest.docx" # if we're creating the report
DEFAULT_GLOSSARY_PATH: str = "./Reports/OrbitalFire-Glossary.csv"
DEFAULT_TECHNICAL_REPORT_PATH: str = "./Reports/OrbitalFire-TechnicalReportDemo.pdf"
DEFAULT_EXECUTIVE_REPORT_PATH: str = "./Reports/OrbitalFire-ExecutiveReportDemo.pdf"
DEFAULT_RECOMMENDATIONS_PATH: str = "./Reports/FindingsDetailsAndRecommendations.xlsx"
SEVERITIES_LEVELS = ["Informational", "Low", "Medium", "High", "Critical"]

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

def activity_log_page_range(activity_report: str) -> str:
    start = None
    last = None
    ACTIVITY_LOG = "Activity Log"
    ACTIVITY_TIME = "Activity Time"
    ACTIVITY_TYPE = "Activity Type"
    ACTIVITY = "Activity"

    arrangement = re.compile(r"\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}")
    with pdfplumber.open(activity_report) as pdf:
        for i, page in enumerate(pdf.pages, start = 1):
            content = page.extract_text() or ""
            separated = content.split()
            cleaned = " ".join(separated)
            if ACTIVITY_LOG in cleaned and ACTIVITY_TIME in cleaned and ACTIVITY_TYPE in cleaned and ACTIVITY in cleaned:
                start = i
            if start is not None and arrangement.search(cleaned):
                last = i
    if start is None or last is None:
        return ""
    page_range = f"{start}-{last}"
    return page_range

def automated_testing_activity(activity_report: str, findings_report: str, user_header: str) -> None:
    page_range = activity_log_page_range(activity_report)
    if page_range == "":
        print("❌ Could not find Activity Log in Activity Report")
        return
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
    doc_holder = Document(findings_report)
    section = "AUTOMATED TESTING ACTIVITY"
    for i, paragraph in enumerate(doc_holder.paragraphs):
        if section.lower() in paragraph.text.lower():
            position = doc_holder.paragraphs[i + 1]
            if user_header.strip():
                header = position.insert_paragraph_before()
                header_run = header.add_run(user_header.strip())
                header_run.font.name = "Corbel"
                header_run.font.size = Pt(12)
                header_run.font.color.rgb = RGBColor(242, 101, 34)
            for entry in activity_log:
                container = position.insert_paragraph_before()
                container.style = "Activity Bullet"
                execute  = container.add_run(entry)
                execute.font.name = "Corbel"
                execute.font.size = Pt(8.5)
                execute.font.color.rgb = RGBColor(64,64,64)
            break
    doc_holder.save(findings_report)
    print("✅ Automated Testing Activity was successfully populated and saved")

def engagement_results_page(executive_report: str) -> str:
    ENGAGEMENT_RESULTS_SUMMARY = "Engagement Results Summary"
    CATEGORY_SUMMARY = "Category Summary"
    PENETRATION_TEST = "Penetration Test"
    with pdfplumber.open(executive_report) as pdf:
        for i, page in enumerate(pdf.pages, start = 1):
            content = page.extract_text() or ""
            separated = content.split()
            cleaned = " ".join(separated)
            if ENGAGEMENT_RESULTS_SUMMARY in cleaned and CATEGORY_SUMMARY in cleaned and PENETRATION_TEST in cleaned:
                page = str(i)
                return page
    return ""

# Populates the Assessment Results Summary section in the Findings Report.
def assessment_results(executive_report: str, findings_report: str) -> None:
    client_call = OpenAI(api_key = os.environ.get("OPENAI_API_KEY"))
    pdf = pymupdf.open(executive_report)
    page = engagement_results_page(executive_report)
    if page == "":
        print("❌ Could not find Engagement Results Summary in Executive Report")
        return
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
        cat_cont2.font.color.rgb = RGBColor(64, 64, 64)
        sum_cont1 = add_pos.insert_paragraph_before()
        sum_cont2 = sum_cont1.add_run(final_sum)
        sum_cont2.font.name = "Corbel"
        sum_cont2.font.color.rgb = RGBColor(64, 64, 64)
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
def recommendations(recommendation_csv: str, technical_report: str, findings_report: str) -> None:
    internal = pd.read_excel(recommendation_csv, sheet_name="Internal")
    external = pd.read_excel(recommendation_csv, sheet_name="External")
    data_container = pd.concat([internal, external], ignore_index=True)
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
        os.makedirs("./Reports/Recommendations", exist_ok=True)
        flame = "./Reports/Recommendations/gray_flame.png"
        flame_holder = body_cont1.add_run()
        flame_holder.add_picture(flame, width=Pt(12))
        body_cont1.add_run(" ")
        body_cont2 = body_cont1.add_run(left_component)
        body_cont2.bold = True
        body_cont2.font.color.rgb = RGBColor(64, 64, 64)
        body_cont2.font.name = "Corbel"
        body_cont2.font.size = Pt(12)
        if right_component:
            normalized_def = body_cont1.add_run(" - " + right_component)
            normalized_def.font.color.rgb = RGBColor(64, 64, 64)
            normalized_def.font.name = "Corbel"
            normalized_def.font.size = Pt(12)
        rec_cont1 = new_position.insert_paragraph_before()
        rec_cont2 = rec_cont1.add_run(f"Refer to {status_holder} Finding {tracker_holder}")
        rec_cont2.font.color.rgb = RGBColor(64, 64, 64)
        rec_cont2.font.name = "Corbel"
        rec_cont2.font.size = Pt(12)
        rec_cont2.underline = True
        rec_cont1.paragraph_format.left_indent = Inches(0.5)
        rec_cont1.paragraph_format.first_line_indent = Inches(0)
        rec_cont1.paragraph_format.space_before = Pt(0)
        rec_cont1.paragraph_format.space_after = Pt(6)
    doc.save(findings_report)
    print("✅ Recommendations was saved.")

# For Appendix. To insert screenshot of image after section header "SCREENSHOT OF OPEN PORTS TABLE"
def add_new_paragraph(paragraph):
    inserted = OxmlElement("w:p")
    paragraph._p.addnext(inserted)
    return Paragraph(inserted, paragraph._parent)

# To populate the Appendix section of the Findings Report.
def appendix(technical_report: str, findings_report: str) -> None:
    title = "Appendix B: Host Discovery (Opened Ports)"
    image_container = "./Reports/Appendix/OPEN_PORTS.PNG"
    os.makedirs("./Reports/Appendix", exist_ok=True)
    pdf_p = pymupdf.open(technical_report)
    target = None
    table_container = None
    len_pdf_p = len(pdf_p)
    for i in range(len_pdf_p):
        page = pdf_p[i]
        text_data = page.get_text()
        if title in text_data:
            target = page
            target_table = page.find_tables()
            if target_table and target_table.tables:
                max_table = max(target_table.tables, key = lambda x: (x.bbox[2] - x.bbox[0]) * (x.bbox[3] - x.bbox[1]))
                x0, y0, x1, y1 = max_table.bbox
                table_container = pymupdf.Rect(x0 - 3, y0 - 3, x1 + 3, y1 + 3)
            else:
                table_header = page.search_for("IP Address")
                protocol = page.search_for("Protocol")
                if protocol and table_header:
                    top = table_header[0].y0 - 3
                    bottom = table_header[0].y1 + 170
                    left_point = table_header[0].x0 - 3
                    right_point = page.rect.width - 36
                    table_container = pymupdf.Rect(left_point, top, right_point, bottom)
                else:
                    table_container = pymupdf.Rect(36, 150, page.rect.width - 36, 320)
            break
    if target is None:
        print(f"The corresponding section: {title} was not found in {technical_report}")
        pdf_p.close()
        return
    image_zoomed = pymupdf.Matrix(2.0, 2.0)
    pixels = target.get_pixmap(matrix=image_zoomed, clip=table_container ,alpha=False)
    pixels.save(image_container)
    pdf_p.close()
    print(f"✅ Screenshot of Open Ports Table was a success.")
    doc = Document(findings_report)
    position = None
    text_1 = "OPEN PORTS |"
    for i, p in enumerate(doc.paragraphs):
        cur_data = p.text.upper()
        if text_1 in cur_data:
            position = i
            break
    if position is None:
        print(f"❌ {text_1} can't be found")
        return
    data_container = doc.paragraphs[position]
    image_struc = add_new_paragraph(data_container)
    image = image_struc.add_run()
    image.add_picture(image_container, width=Inches(6.5))
    image_struc.paragraph_format.space_before = Pt(0)
    image_struc.paragraph_format.space_after = Pt(6)
    print(f"✅ Open Ports Table was pasted over Findings Report successfully.")
    doc.save(findings_report)
    print("✅ Findings Report was successfully saved.")

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

def delete_page_brk_before(p):
    p.paragraph_format.page_break_before = False
    marker = ".//w:br[@w:type='page']"
    for brk in p._element.xpath(marker):
        ancestor = brk.getparent()
        if ancestor is not None:
            ancestor.remove(brk)

def delete_sec_break_from_prev(p):
    prev = p._element.getprevious()
    if prev is None:
        return
    marker = ".//w:sectPr"
    previous_section = prev.xpath(marker)
    for section in previous_section:
        ancestor = section.getparent()
        if ancestor is not None:
            ancestor.remove(section)

def resize_image(p, file_image: str, max_w: float = 6.3, max_h: float = 3.8, desired_ra: float = 3.8) -> None:
    with Image.open(file_image) as image:
        pixel_w, pixel_h = image.size
        if pixel_w == 0 or pixel_h == 0:
            return
        prop = pixel_w / pixel_h
        modified_w = max_w
        if prop > desired_ra:
            modified_w = 5.8
        final_w = modified_w
        final_h = final_w / prop
        if max_h < final_h:
            final_h = max_h
            final_w = final_h * prop
        p.add_run().add_picture(file_image, width=Inches(final_w), height=Inches(final_h))

def modify_bounds(container, top, bottom, left, right, x_padding = 4, y_padding = 4):
    max_left = max(left, container.x0 - x_padding)
    max_top = max(top, container.y0 - y_padding)
    min_right = min(right, container.x1 + x_padding)
    min_bottom = min(bottom, container.y1 + y_padding)
    result = pymupdf.Rect(max_left, max_top, min_right, min_bottom)
    return result

def center(container, top, bottom, left, right, extra_w = 18):
    center_data = (container.x0 + container.x1) / 2
    target_w = min(right - left, container.width + extra_w)
    alter_x0 = center_data - (target_w / 2)
    alter_x1 = center_data + (target_w / 2)
    if left > alter_x0:
        move = left - alter_x0
        alter_x0 += move
        alter_x1 += move
    if right < alter_x1:
        move = alter_x1 - right
        alter_x0 -= move
        alter_x1 -= move
    alter_x0 = max(left, alter_x0)
    alter_x1 = min(right, alter_x1)
    mod_top = max(top, container.y0 - 3)
    mod_bottom = min(bottom, container.y1 + 3)
    res = pymupdf.Rect(alter_x0, mod_top, alter_x1, mod_bottom)
    return res

def construct_portion(p, top, bottom, table_type="text"):
    left_side = 36
    right_side = p.rect.width - 36
    table_collector = []
    alternative_table = []
    if table_type == "table":
        try:
            detected_table = p.find_tables()
            if detected_table and detected_table.tables:
                for table in detected_table.tables:
                    x0, y0, x1, y1 = table.bbox
                    holder = pymupdf.Rect(x0, y0, x1, y1)
                    if holder.y0 >= (top - 10) and holder.y1 <= (bottom + 12):
                        if holder.width >= 40 and holder.height >= 8:
                            alternative_table.append(holder)
            if alternative_table:
                optimal = min(alternative_table, key=lambda x: abs(x.y0 - top))
                optimal = modify_bounds(optimal, top, bottom, left_side, right_side, x_padding=4, y_padding=4)
                optimal = center(optimal, top, bottom, left_side, right_side, extra_w=10)
                return optimal
        except Exception as e:
            pass
    d_boxes = []
    try:
      draws = p.get_drawings()
      for d in draws:
          temp = d.get("rect")
          if not temp:
              continue
          if temp.y0 >= (top - 10) and temp.y1 <= (bottom + 12):
              if temp.width >= 40 and temp.height >= 8:
                  d_boxes.append(temp)
    except Exception as e:
        d_boxes = []
    if d_boxes:
        optimal = min(d_boxes, key=lambda x: abs(x.y0 - top))
        combine = pymupdf.Rect(optimal.x0, optimal.y0, optimal.x1, optimal.y1)
        marked = True
        while marked:
            marked = False
            for entry in d_boxes:
                h_matching = False
                vertical_mark = False
                overlap = False
                if abs(entry.x0 - combine.x0) < 30 and abs(entry.x1 - combine.x1) < 30:
                    h_matching = True
                if -8 <= (entry.y0 - combine.y1) <= 20:
                    vertical_mark = True
                if not ((entry.y1 < combine.y0 - 3) or (entry.y0 > combine.y1 + 3)):
                    overlap = True
                if h_matching and (vertical_mark or overlap):
                    new_x0 = min(combine.x0, entry.x0)
                    new_y0 = min(combine.y0, entry.y0)
                    new_x1 = max(combine.x1, entry.x1)
                    new_y1 = max(combine.y1, entry.y1)
                    res = pymupdf.Rect(new_x0, new_y0, new_x1, new_y1)
                    if res != combine:
                        combine = res
                        marked = True
        if table_type == "text":
            modified_bot = min(bottom, combine.y1 + 8)
            combine = pymupdf.Rect(combine.x0, combine.y0, combine.x1, modified_bot)
            contain = modify_bounds(combine, top, bottom, left_side, right_side, x_padding=4, y_padding=4)
            return contain
        combine = modify_bounds(combine, top, bottom, left_side, right_side, x_padding=4, y_padding=4)
        combine = center(combine, top, bottom, left_side, right_side, extra_w=10)
        return combine
    text = p.get_text("words")
    for i in text:
        x0, y0, x1, y1, container = i[:5]
        container = str(container).strip()
        if not container:
            continue
        if y1 <= bottom and y0 >= top:
            res = pymupdf.Rect(x0, y0, x1, y1)
            table_collector.append(res)
    if table_collector:
        x0_min = table_collector[0].x0
        y0_min = table_collector[0].y0
        x1_max = table_collector[0].x1
        y1_max = table_collector[0].y1
        for row in table_collector:
            if row.x0 < x0_min:
                x0_min = row.x0
            if row.y0 < y0_min:
                y0_min = row.y0
            if row.x1 > x1_max:
                x1_max = row.x1
            if row.y1 > y1_max:
                y1_max = row.y1
        combine = pymupdf.Rect(x0_min, y0_min, x1_max, y1_max)
        if table_type == "text":
            mod_bottom = min(bottom, combine.y1 + 10)
            combine = pymupdf.Rect(combine.x0, combine.y0, combine.x1, mod_bottom)
            output = modify_bounds(combine, top, bottom, left_side, right_side, x_padding=4, y_padding=4)
            return output
        combine = modify_bounds(combine, top, bottom, left_side, right_side, x_padding=4, y_padding=4)
        combine = center(combine, top, bottom, left_side, right_side, extra_w=10)
        return combine
    default = pymupdf.Rect(left_side, top, right_side, bottom)
    return default

def enlarge_pic(box, container, left_padding=2, right_padding=10, y_padding_top=6, y_padding_bottom=10):
    max_x0 = max(container.x0, box.x0 - left_padding)
    max_y0 = max(container.y0, box.y0 - y_padding_top)
    min_x1 = min(container.x1, box.x1 + right_padding)
    min_y1 = min(container.y1, box.y1 + y_padding_bottom)
    res = pymupdf.Rect(max_x0, max_y0, min_x1, min_y1)
    return res

def left_trim(container, unit=8):
    res = pymupdf.Rect(container.x0 + unit, container.y0, container.x1, container.y1)
    return res

def bottom_trim(container, unit=8):
    res = pymupdf.Rect(container.x0, container.y0, container.x1, container.y1 - unit)
    return res

def data_is_present(page, top, bottom,) -> bool:
    target = []
    discard = {"ORBITALFIRE", "PENETRATION", "TESTING"}
    struc = pymupdf.Rect(36, top, page.rect.width - 36, bottom)
    content = page.get_text("words", clip=struc)
    for entry in content:
        data = str(entry[4]).strip()
        if not data:
            continue
        if data.upper() in discard:
            continue
        if data.isdigit():
            continue
        target.append(data)
    result = " ".join(target).strip()
    length = len(result)
    if length >= 5:
        return True
    return False

def finding_details(technical_path: str, findings_report: str, details_path: str, customer_name: str) -> None:
    internal = pd.read_excel(details_path, sheet_name="Internal")
    external = pd.read_excel(details_path, sheet_name="External")
    dataframe = pd.concat([internal, external], ignore_index=True)
    columns = []
    for i in dataframe.columns:
        clean = str(i).strip()
        columns.append(clean)
    dataframe.columns = columns
    find_mapping = {}
    for _, i in dataframe.iterrows():
        title_holder = str(i.get("Finding Title", "")).strip()
        description_holder = str(i.get("Finding Description", "")).strip()
        risk_holder = str(i.get("Risk Level", "")).strip()
        if title_holder.lower() == "nan" or title_holder == "":
            continue
        if description_holder.lower() == "nan" or description_holder == "":
            continue
        if risk_holder.lower() == "nan" or risk_holder == "":
            continue
        new_title = " ".join(title_holder.lower().replace("\n" ," ").replace("–", "-"). replace("—", "-").split()).strip()
        find_mapping[new_title] = {"description": description_holder, "risk": risk_holder.title()}
    pdf = pymupdf.open(technical_path)
    severities = ["CRITICAL", "HIGH" ,"MEDIUM", "LOW"]
    finds = []
    viewed_titles = set()
    len_pdf = len(pdf)
    for page_number in range(len_pdf):
        p = pdf[page_number]
        content = p.get_text("text")
        lines = []
        for i in content.splitlines():
            clean = i.strip()
            if clean:
                lines.append(clean)
        counter = 0
        length = len(lines)
        while counter < length:
            line = lines[counter]
            matching = re.match(r"^(CRITICAL|HIGH|MEDIUM|LOW)\s+(.+)$", line, re.IGNORECASE)
            if matching:
                SEVERITY = matching.group(1).title()
                title = matching.group(2).strip()
                cleaned_title = " ".join(title.lower().replace("\n" ," ").replace("–", "-").replace("—", "-").split()).strip()
                page_content_u = content.upper()
                detail_p = ("AFFECTED NODES" in page_content_u or "OBSERVATION" in page_content_u or "RECOMMENDATION" in page_content_u or "EVIDENCE" in page_content_u)
                if cleaned_title in find_mapping and cleaned_title not in viewed_titles and detail_p:
                    viewed_titles.add(cleaned_title)
                    finds.append({"SEVERITY": find_mapping[cleaned_title]["risk"], "title": title, "description": find_mapping[cleaned_title]["description"], "page": page_number})
                counter += 1
                continue
            if line.upper() in severities and (counter + 1) < length:
                SEVERITY = line.title()
                title = lines[counter + 1].strip()
                cleaned_title = " ".join(title.lower().replace("\n" ," ").replace("–", "-").replace("—", "-").split()).strip()
                page_content_u = content.upper()
                detail_p = ("AFFECTED NODES" in page_content_u or "OBSERVATION" in page_content_u or "RECOMMENDATION" in page_content_u or "EVIDENCE" in page_content_u)
                if cleaned_title in find_mapping and cleaned_title not in viewed_titles and detail_p:
                    viewed_titles.add(cleaned_title)
                    finds.append({"SEVERITY": find_mapping[cleaned_title]["risk"], "title": title, "description": find_mapping[cleaned_title]["description"], "page": page_number})
                counter += 2
                continue
            counter += 1
    image_directory = "./Reports/Findings_Details"
    os.makedirs(image_directory, exist_ok=True)
    for i, find in enumerate(finds, start=1):
        find["affected_nodes_table"] = []
        find["evidence_table"] = []
        starting = find["page"]
        length_pdf = len(pdf)
        up_bound = min(starting + 10, length_pdf)
        for x in range(starting, up_bound):
            page = pdf[x]
            if not find["affected_nodes_table"]:
                affected = page.search_for("Affected Nodes")
                if not affected:
                    affected = page.search_for("AFFECTED NODES")
                if affected:
                    top = affected[0].y1 + 3
                    bottom = page.rect.height - 36
                    holder_val = ["Recommendation", "RECOMMENDATION",
                                  "Reproduction Steps", "REPRODUCTION STEPS", "References", "REFERENCES",
                                  "Evidence", "EVIDENCE"] + severities
                    connections = []
                    for entry in holder_val:
                        collect = page.search_for(entry)
                        for k in collect:
                            if k.y0 > top:
                                result = k.y0 - 4
                                connections.append(result)
                    if connections:
                        bottom = min(connections)
                    if top < bottom:
                        crop = construct_portion(page, top, bottom, table_type="table")
                        section_cont = pymupdf.Rect(36, top, page.rect.width - 36, bottom)
                        crop = enlarge_pic(crop, section_cont, left_padding=2, right_padding=10, y_padding_top=10, y_padding_bottom=10)
                        crop = left_trim(crop, unit=12)
                        crop = bottom_trim(crop, unit=2)
                        file_path_affected = os.path.join(image_directory, f"Affected_Sample_{i}_part1.png")
                        page.get_pixmap(matrix=pymupdf.Matrix(2,2), clip=crop, alpha=False).save(file_path_affected)
                        find["affected_nodes_table"].append(file_path_affected)
                        if not connections:
                            to_stop = False
                            counter_part = 2
                            follow_up = x + 1
                            for next_num in range(follow_up, up_bound):
                                ending_pts = []
                                next = pdf[next_num]
                                continue_top = 36
                                continue_bottom = next.rect.height - 36
                                stopper = ["Recommendation", "RECOMMENDATION",
                                           "Reproduction Steps", "REPRODUCTION STEPS", "References", "REFERENCES",
                                           "Evidence", "EVIDENCE"] + severities
                                for stop in stopper:
                                    for b in next.search_for(stop):
                                        if continue_top < b.y0:
                                            reduced_space = b.y0 - 4
                                            ending_pts.append(reduced_space)
                                if ending_pts:
                                    to_stop = True
                                    continue_bottom = min(ending_pts)
                                if continue_bottom > continue_top:
                                    continue_crop = construct_portion(next, continue_top, continue_bottom, table_type="table")
                                    continue_section = pymupdf.Rect(36, continue_top, next.rect.width - 36, continue_bottom)
                                    continue_crop = enlarge_pic(continue_crop, continue_section, left_padding=2, right_padding=10, y_padding_top=10, y_padding_bottom=10)
                                    continue_crop = left_trim(continue_crop, unit=12)
                                    continue_crop = bottom_trim(continue_crop, unit=2)
                                    file_path_affected_continue = os.path.join(image_directory, f"Affected_Sample_{i}_part{counter_part}.png")
                                    next.get_pixmap(matrix=pymupdf.Matrix(2,2), clip=continue_crop, alpha=False).save(file_path_affected_continue)
                                    find["affected_nodes_table"].append(file_path_affected_continue)
                                    counter_part += 1
                                if to_stop:
                                    break
            if not find["evidence_table"]:
                evidence = page.search_for("Evidence")
                if not evidence:
                    evidence = page.search_for("EVIDENCE")
                if evidence:
                    evidence_box = max(evidence, key=lambda s: s.y0)
                    top = evidence_box.y1 + 3
                    bottom = page.rect.height - 36
                    connections = []
                    FLAG = severities + ["References", "REFERENCES", "Recommendation", "RECOMMENDATION",
                                         "Reproduction Steps", "REPRODUCTION STEPS", "Appendix", "APPENDIX"]
                    for marker in FLAG:
                        container = page.search_for(marker)
                        for k in container:
                            if k.y0 > top:
                                result = k.y0
                                connections.append(result)
                    if connections:
                        bottom = min(connections)
                    if top < bottom:
                        crop = construct_portion(page, top, bottom, table_type="text")
                        section_cont = pymupdf.Rect(36, top, page.rect.width - 36, bottom)
                        crop = enlarge_pic(crop, section_cont, left_padding=2, right_padding=6, y_padding_top=4, y_padding_bottom=6)
                        crop = bottom_trim(crop, unit=10)
                        file_path_evidence = os.path.join(image_directory, f"Evidence_Sample_{i}_part1.png")
                        page.get_pixmap(matrix=pymupdf.Matrix(2,2), clip=crop, alpha=False).save(file_path_evidence)
                        find["evidence_table"].append(file_path_evidence)
                        if not connections:
                            to_stop = False
                            counter_part = 2
                            follow_up = x + 1
                            for next_num in range(follow_up, up_bound):
                                ending_pts = []
                                next = pdf[next_num]
                                continue_top = 36
                                continue_bottom = next.rect.height - 36
                                stopper = severities + ["References", "REFERENCES", "Recommendation", "RECOMMENDATION",
                                                        "Reproduction Steps", "REPRODUCTION STEPS", "Appendix", "APPENDIX"]
                                for stop in stopper:
                                    for b in next.search_for(stop):
                                        if continue_top < b.y0:
                                            reduced_space = b.y0 - 4
                                            ending_pts.append(reduced_space)
                                if ending_pts:
                                    to_stop = True
                                    continue_bottom = min(ending_pts)
                                is_present = data_is_present(next, continue_top, continue_bottom)
                                if not is_present:
                                    if to_stop:
                                        break
                                    continue
                                if continue_bottom > continue_top:
                                    continue_crop = construct_portion(next, continue_top, continue_bottom, table_type="text")
                                    continue_section = pymupdf.Rect(36, continue_top, next.rect.width - 36, continue_bottom)
                                    continue_crop = enlarge_pic(continue_crop, continue_section, left_padding=2, right_padding=10, y_padding_top=10, y_padding_bottom=10)
                                    continue_crop = left_trim(continue_crop, unit=12)
                                    continue_crop = bottom_trim(continue_crop, unit=2)
                                    file_path_evidence_continue = os.path.join(image_directory, f"Evidence_Sample_{i}_part{counter_part}.png")
                                    next.get_pixmap(matrix=pymupdf.Matrix(2,2), clip=continue_crop, alpha=False).save(file_path_evidence_continue)
                                    find["evidence_table"].append(file_path_evidence_continue)
                                    counter_part += 1
                                if to_stop:
                                    break
    pdf.close()
    doc = Document(findings_report)
    STATUS = {"Critical": [], "High": [], "Medium": [], "Low": []}
    for entry in finds:
        if entry["SEVERITY"] in STATUS:
            content = entry["SEVERITY"]
            STATUS[content].append(entry)
    collector_status = ["Critical", "High", "Medium", "Low"]
    for entry in collector_status:
        cur_finds = STATUS[entry]
        position = None
        move_index = None
        for i, paragraph in enumerate(doc.paragraphs):
            if paragraph.text.strip().upper() == entry.upper():
                position = paragraph
                move_index = i
                break
        if position is None:
            print(f"Severity header {entry} can't found")
            continue
        if not cur_finds:
            continue
        current_index = collector_status.index(entry)
        prev_has_info = False
        for prev in collector_status[:current_index]:
            info_exist = len(STATUS[prev])
            if info_exist > 0:
                prev_has_info = True
                break
        if prev_has_info:
            position.paragraph_format.page_break_before = True
        else:
            delete_page_brk_before(position)
            delete_sec_break_from_prev(position)
        modified_index = move_index + 1
        length_doc_p = len(doc.paragraphs)
        upper = min(move_index + 5, length_doc_p)
        for sample in range(modified_index, upper):
            data = doc.paragraphs[sample].text.strip().lower()
            EMPTY = ["none", "• none", "o none", "· none"]
            if data in EMPTY:
                element = doc.paragraphs[sample]._element
                element.getparent().remove(element)
                break
        insert_after = position
        counting = 1
        for i in cur_finds:
            para1 = add_new_paragraph(insert_after)
            run_para0 = para1.add_run(f"{counting}.    ")
            run_para0.font.color.rgb = RGBColor(64, 64, 64)
            run_para0.font.name = "Corbel"
            run_para0.font.size = Pt(12)
            run_para1 = para1.add_run(f"{i['title']}")
            run_para1.bold = True
            run_para1.font.color.rgb = RGBColor(64, 64, 64)
            run_para1.font.name = "Corbel"
            run_para1.font.size = Pt(12)
            run_para_hyphen = para1.add_run(" - ")
            run_para_hyphen.font.color.rgb = RGBColor(64, 64, 64)
            run_para_hyphen.font.name = "Corbel"
            run_para_hyphen.font.size = Pt(12)
            description_modified = i["description"]
            if customer_name.strip():
                description_modified = description_modified.replace("[CUSTOMER NAME]", customer_name.strip())
            run_para2 = para1.add_run(description_modified)
            run_para2.font.color.rgb = RGBColor(64, 64, 64)
            run_para2.font.name = "Corbel"
            run_para2.font.size = Pt(12)
            insert_after = para1
            counting += 1
            if i["affected_nodes_table"]:
                para_an = add_new_paragraph(insert_after)
                run_para_an = para_an.add_run("Affected Nodes:")
                run_para_an.font.italic = True
                run_para_an.font.color.rgb = RGBColor(64, 64, 64)
                run_para_an.font.name = "Corbel"
                run_para_an.font.size = Pt(12)
                insert_after = para_an
                for affected_image in i["affected_nodes_table"]:
                    if os.path.exists(affected_image):
                        para2 = add_new_paragraph(insert_after)
                        resize_image(para2, affected_image, max_w=6.3, max_h=2.8, desired_ra=3.8)
                        insert_after = para2
            para3 = add_new_paragraph(insert_after)
            run_para3 = para3.add_run("Evidence:")
            run_para3.italic = True
            run_para3.font.color.rgb = RGBColor(64, 64, 64)
            run_para3.font.name = "Corbel"
            run_para3.font.size = Pt(12)
            insert_after = para3
            if i["evidence_table"]:
                for evidence_image in i["evidence_table"]:
                    if os.path.exists(evidence_image):
                        para4 = add_new_paragraph(insert_after)
                        resize_image(para4, evidence_image, max_w=6.3, max_h=3.8, desired_ra=3.8)
                        insert_after = para4
    doc.save(findings_report)
    print(f"✅ Findings Details was successfully populated and saved")

def IPAddress(technical: str, findings_report_path: str) -> None:
    IP_Address_sec = False
    table_head = "IP ADDRESSES & RANGES"
    Area = False
    collector = []
    engage_sec = "Engagement Scope of Work"
    with pdfplumber.open(technical) as pdf:
        for page in pdf.pages:
            content = page.extract_text()
            if not content:
                continue
            entries = content.splitlines()
            for entry in entries:
                split_up = entry.split()
                cleaned = " ".join(split_up).strip()
                if cleaned == "":
                    continue
                if engage_sec in cleaned:
                    Area = True
                    continue
                if not Area:
                    continue
                if table_head in cleaned:
                    IP_Address_sec = True
                    continue
                if IP_Address_sec and ("Agent Information" in cleaned or "Task Performed" in cleaned or "Rules of Engagement" in cleaned):
                    IP_Address_sec = False
                    break
                if IP_Address_sec:
                    discovered_ip = re.findall(r"\b(?:\d{1,3}\.){3}\d{1,3}(?:/\d{1,2})?\b", cleaned)
                    for IP in discovered_ip:
                        if IP not in collector:
                            collector.append(IP)
        if not collector:
            print("❌ Could not find IP Addresses in Technical Report")
            return
        print("✅ IP Address was successfully parsed from Technical Report")
        doc = Document(findings_report_path)
        position = None
        target = "The following IP Addresses and hosts were evaluated during automated testing:"
        for i, p in enumerate(doc.paragraphs):
            if target.lower() in p.text.lower():
                position = i
                break
        if position is None:
            print("❌ Could not find IP Address section in Findings Report")
            return
        increment = position + 1
        index = doc.paragraphs[increment]
        for IP in collector:
            para = index.insert_paragraph_before()
            para.style = "IP"
            execute = para.add_run(IP)
            execute.font.name = "Corbel"
            execute.font.size = Pt(12)
            execute.font.color.rgb = RGBColor(64, 64, 64)
        doc.save(findings_report_path)
        print("✅ IP Address was successfully populated and saved")

def place_host_data(paragraph, temp: str, data: str) -> bool:
    for entry in paragraph.runs:
        if temp in entry.text:
            entry.text = entry.text.replace(temp, data)
            return True
    return False

def is_singular_plural(value: str) -> bool:
    processed = value.strip().lower()
    if processed == "one" or processed == "1" or processed == "(1)":
        return True
    if re.search(r"\(\s*1\s*\)", processed):
        return True
    return False

def Host_Discovery(technical: str, findings_report_path: str) -> None:
    start_point = "Host Discovery"
    end_point = "Enumeration"
    Host_Discovery_sec = False
    network_ranges = None
    systems = None
    addresses = None
    open_ports = None
    target_1 = "IP address/range that was provided as part of the scope"
    target_2 = "address/range that was scanned"
    NETWORKS_HOLDER = "[NETWORK RANGE]"
    NETWORK_RANGE_WORD = "[NETWORK RANGE WORD]"
    NETWORK_RANGE_VERB = "[NETWORK RANGE VERB]"
    SYSTEMS_HOLDER = "[SYSTEMS]"
    SYSTEM_WORD = "[SYSTEM WORD]"
    ADDRESSES_HOLDER = "[ADDRESSES]"
    ADDRESS_WORD = "[ADDRESS WORD]"
    ADDRESS_VERB = "[ADDRESS VERB]"
    OPEN_PORTS_HOLDER = "[OPEN PORTS]"
    OPEN_PORT_WORD = "[OPEN PORT WORD]"
    with pdfplumber.open(technical) as pdf:
        for page in pdf.pages:
            data = page.extract_text()
            if not data:
                continue
            entries = data.splitlines()
            for entry in entries:
                split_up = entry.split()
                cleaned = " ".join(split_up).strip()
                if cleaned == "":
                    continue
                if cleaned == start_point:
                    Host_Discovery_sec = True
                    continue
                if Host_Discovery_sec and cleaned == end_point:
                    Host_Discovery_sec = False
                    break
                if not Host_Discovery_sec:
                    continue
                if target_1 in cleaned:
                    first_match = re.search(r"Of the\s+(.+?)\s+IP address/range.*?identify a total of\s+(.+?)\s+systems?", cleaned, re.IGNORECASE)
                    if first_match:
                        network_ranges = first_match.group(1).strip()
                        systems = first_match.group(2).strip()
                if target_2 in cleaned:
                    second_match = re.search(r"Of the\s+(.+?)\s+address/range.*?(?:found|identified)\s+(.+?)\s+(?:ports opened|open ports)", cleaned, re.IGNORECASE)
                    if second_match:
                        addresses = second_match.group(1).strip()
                        open_ports = second_match.group(2).strip()
    if not all([network_ranges, systems, addresses, open_ports]):
        print("❌ Could not do data extraction from Host Discovery in Technical Report")
        return
    if is_singular_plural(network_ranges):
        network_word = "network range"
        network_verb = "was"
    else:
        network_word = "network ranges"
        network_verb = "were"
    if is_singular_plural(systems):
        system_word = "system"
    else:
        system_word = "systems"
    if is_singular_plural(addresses):
        address_word = "address"
        address_verb = "was"
    else:
        address_word = "addresses"
        address_verb = "were"
    if is_singular_plural(open_ports):
        open_port_word = "open port"
    else:
        open_port_word = "open ports"
    doc = Document(findings_report_path)
    position = None
    for p in doc.paragraphs:
        para = p.text
        if NETWORKS_HOLDER in para or SYSTEMS_HOLDER in para or ADDRESSES_HOLDER in para or OPEN_PORTS_HOLDER in para:
            position = p
            break
    if position is None:
        print("❌ Could not find Host Discovery section in Findings Report")
        return
    new_info1 = place_host_data(position, NETWORKS_HOLDER, network_ranges)
    new_info2 = place_host_data(position, SYSTEMS_HOLDER, systems)
    new_info3 = place_host_data(position, ADDRESSES_HOLDER, addresses)
    new_info4 = place_host_data(position, OPEN_PORTS_HOLDER, open_ports)
    new_info5 = place_host_data(position, NETWORK_RANGE_WORD, network_word)
    new_info6 = place_host_data(position, NETWORK_RANGE_VERB, network_verb)
    new_info7 = place_host_data(position, SYSTEM_WORD, system_word)
    new_info8 = place_host_data(position, ADDRESS_WORD, address_word)
    new_info9 = place_host_data(position, ADDRESS_VERB, address_verb)
    new_info10 = place_host_data(position, OPEN_PORT_WORD, open_port_word)
    if not all([new_info1, new_info2, new_info3, new_info4, new_info5, new_info6, new_info7, new_info8, new_info9, new_info10]):
        print("❌ Could not replace one of the placeholder values under Host Discovery in Findings Report")
        return
    doc.save(findings_report_path)
    print("✅ Host Discovery was successfully populated and saved")

def normalize_anchor_text(text: str) -> str:
    return " ".join(text.upper().split()).strip()

def set_para(p) -> None:
    for entry in p.runs:
        entry.font.name = "Corbel"
        entry.font.size = Pt(12)
        entry.font.color.rgb = RGBColor(64, 64, 64)

def indent_after_insert(body, destination) -> None:
    destination.paragraph_format.left_indent = Inches(0.25)
    destination.paragraph_format.right_indent = body.paragraph_format.right_indent
    destination.paragraph_format.first_line_indent = Inches(0)


def insert_image_under_heading(doc: Document, heading_text: str, image_path: str):
    """
    Inserts an image under a heading anchor in the document.
    Returns the image paragraph on success, None on failure.
    """
    normalized_heading = normalize_anchor_text(heading_text)
    target_paragraph = None

    #print(f"🔍 Searching for anchor: '{heading_text}' (normalized: '{normalized_heading}')")

    # First pass: exact match
    for p in doc.paragraphs:
        if normalize_anchor_text(p.text) == normalized_heading:
            target_paragraph = p
            #print(f"✓ Found exact match at paragraph: '{p.text[:80]}'")
            break

    # Second pass: paragraph starts with the anchor text (e.g., "SUB DOMAIN TABLE:" or "SUB DOMAIN TABLE -")
    if target_paragraph is None:
        for p in doc.paragraphs:
            normalized_p_text = normalize_anchor_text(p.text)
            if normalized_p_text.startswith(normalized_heading):
                target_paragraph = p
                #print(f"✓ Found startswith match at paragraph: '{p.text[:80]}'")
                break

    # Third pass: the anchor is the ONLY significant text in the paragraph (allowing for punctuation)
    if target_paragraph is None:
        for p in doc.paragraphs:
            normalized_p_text = normalize_anchor_text(p.text)
            # Remove common punctuation and check if what's left is just the anchor
            cleaned = normalized_p_text.strip(":-–—.,;!?()[]{}\"'")
            if cleaned == normalized_heading:
                target_paragraph = p
                #print(f"✓ Found cleaned match at paragraph: '{p.text[:80]}'")
                break

    # Fourth pass: anchor text is contained anywhere within the paragraph (e.g., "Reference:\nDOPPELGANGER DOMAINS TABLE")
    if target_paragraph is None:
        for p in doc.paragraphs:
            normalized_p_text = normalize_anchor_text(p.text)
            if normalized_heading in normalized_p_text:
                target_paragraph = p
                #print(f"✓ Found anchor embedded in paragraph: '{p.text[:100]}...'")
                break

    if target_paragraph is None:
        #print(f"✗ Anchor '{heading_text}' not found in document")
        #print(f"  📄 Showing paragraphs containing relevant keywords:")
        search_terms = ['DOPPELGANGER', 'SUB DOMAIN', 'DNS RECORD', 'REFERENCE', 'TABLE', 'DOMAIN INFORMATION']
        #for idx, p in enumerate(doc.paragraphs):
            #if p.text.strip():
                #p_upper = p.text.upper()
                #if any(term in p_upper for term in search_terms):
                    #print(f"    [{idx}] '{p.text[:200]}'")
        return None

    # Split the paragraph at the anchor text position
    original_text = target_paragraph.text

    # Find anchor position (case-insensitive)
    anchor_pattern = re.compile(re.escape(heading_text), re.IGNORECASE)
    match = anchor_pattern.search(original_text)

    if not match:
        print(f"❌ Anchor text matching failed (shouldn't happen)")
        return None

    text_before = original_text[:match.start()].rstrip()
    text_after = original_text[match.end():].lstrip()

    #print(f"  📝 Original paragraph length: {len(original_text)} chars")
    #print(f"  📝 Text before anchor: {len(text_before)} chars")
    #print(f"  📝 Anchor text: '{heading_text}'")
    #print(f"  📝 Text after anchor: {len(text_after)} chars")
    #print(f"  → Image file exists: {os.path.exists(image_path)}")

    # Update the original paragraph to contain only text before the anchor
    target_paragraph.text = text_before
    set_para(target_paragraph)
    # Create paragraph for the image
    image_para = add_new_paragraph(target_paragraph)
    indent_after_insert(target_paragraph, image_para)
    inserted_ok = resize_image_modified(image_para, image_path, max_w=6.3, max_h=3.8, desired_ra=3.8)

    if not inserted_ok:
        # Rollback: restore original text and remove image paragraph
        target_paragraph.text = original_text
        parent = image_para._element.getparent()
        if parent is not None:
            parent.remove(image_para._element)
        print(f"❌ Image insertion failed for '{heading_text}' - changes rolled back")
        return None

    image_para.paragraph_format.space_before = Pt(6)
    image_para.paragraph_format.space_after = Pt(6)

    # If there's text after the anchor, create a new paragraph for it
    if text_after:
        after_para = add_new_paragraph(image_para)
        indent_after_insert(target_paragraph, after_para)
        after_para.text = text_after
        set_para(after_para)
        after_para.paragraph_format.space_before = Pt(6)
        #print(f"  ✓ Split paragraph: kept {len(text_before)} chars before, inserted image, moved {len(text_after)} chars after")
    #else:
        #print(f"  ✓ Replaced anchor with image (no text after)")

    return image_para

def resize_image_modified(p, file_image: str, max_w: float = 6.3, max_h: float = 3.8, desired_ra: float = 3.8) -> bool:
    if not os.path.exists(file_image):
        print(f"WARNING: Image file not found for insertion: {file_image}")
        return False

    try:
        with Image.open(file_image) as image:
            pixel_w, pixel_h = image.size
            if pixel_w == 0 or pixel_h == 0:
                return False
            prop = pixel_w / pixel_h
            modified_w = max_w
            if prop > desired_ra:
                modified_w = 5.8
            final_w = modified_w
            final_h = final_w / prop
            if max_h < final_h:
                final_h = max_h
                final_w = final_h * prop
            p.add_run().add_picture(file_image, width=Inches(final_w), height=Inches(final_h))
        return True
    except Exception as e:
        print(f"WARNING: Failed to insert image '{file_image}': {e}")
        return False

def capture_dns_continuation_parts(pdf, start_page_index: int, second_table_rect: pymupdf.Rect, output_path: str, image_zoomed: pymupdf.Matrix) -> tuple:
    """
    Captures a DNS table that may be split across multiple pages.
    HARDCODED: Always captures the next page as continuation.
    Returns tuple of (first_image_path, second_image_path) or (first_image_path, None) if no continuation.
    """
    #print(f"  📊 Capturing DNS table starting on page {start_page_index + 1}...")

    start_page = pdf[start_page_index]

    # Capture the first part of the DNS table
    first_pixels = start_page.get_pixmap(matrix=image_zoomed, clip=second_table_rect, alpha=False)
    first_pixels.save(output_path)

    # Apply 25% bottom crop to remove any page footer/noise
    with Image.open(output_path) as img:
        width, height = img.size
        trimmed_height = int(height * 0.75)
        if trimmed_height > 0 and trimmed_height < height:
            img.crop((0, 0, width, trimmed_height)).save(output_path)
            #print(f"    ✂️  Cropped bottom 25% from first DNS capture")

    # HARDCODED: Always capture the immediate next page as continuation
    next_page_index = start_page_index + 1

    if next_page_index >= len(pdf):
        #print(f"    ℹ️  No next page available - DNS table on last page")
        return (output_path, None)

    #print(f"    📄 Capturing continuation from page {next_page_index + 1} (HARDCODED)...")

    next_page = pdf[next_page_index]

    # Start capture from top of page, minimal offset to avoid cutting off rows
    # HARDCODED: Capture top portion of page 8 (table continuation area only)
    continuation_top = 36
    # Capture approximately top 2/3 of page to get full table without excessive whitespace
    continuation_bottom = next_page.rect.height * 0.65
    continuation_clip = pymupdf.Rect(36, continuation_top, next_page.rect.width - 36, continuation_bottom)
    continuation_path = "./Reports/Informational/DNS-Record2.png"
    next_page.get_pixmap(matrix=image_zoomed, clip=continuation_clip, alpha=False).save(continuation_path)

    # Apply cropping to remove ALL surrounding whitespace (match other tables)
    with Image.open(continuation_path) as img:
        width, height = img.size
        # Aggressive crop: remove whitespace from all sides to match other table style
        crop_left = int(width * 0.05)   # Remove left margin
        crop_top = int(height * 0.02)   # Minimal top to keep first row
        crop_right = int(width * 0.98)  # Remove right margin
        crop_bottom = int(height * 0.45)  # Keep only table rows, remove bottom whitespace
        if crop_bottom > crop_top and crop_right > crop_left:
            img.crop((crop_left, crop_top, crop_right, crop_bottom)).save(continuation_path)
            #print(f"      ✂️  Cropped all whitespace from continuation table")

    #print(f"      ✅ Captured continuation to {continuation_path}")

    return (output_path, continuation_path)

def number_to_word(n: int) -> str:
    """Convert number to word form (1-20 supported)"""
    words = {
        1: "one", 2: "two", 3: "three", 4: "four", 5: "five",
        6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten",
        11: "eleven", 12: "twelve", 13: "thirteen", 14: "fourteen", 15: "fifteen",
        16: "sixteen", 17: "seventeen", 18: "eighteen", 19: "nineteen", 20: "twenty"
    }
    return words.get(n, str(n))

def Informational(technical_report: str, findings_report_path: str) -> None:
    pdf = pymupdf.open(technical_report)
    os.makedirs("./Reports/Informational", exist_ok=True)

    doppelganger_image = None
    subdomain_image = None
    dns_image = None
    first_table_candidates = []
    doppelganger_count = None
    doppelganger_page_index = None

    for i in range(len(pdf)):
        page = pdf[i]
        ev = page.search_for("Evidence")
        inf = page.search_for("Informational")

        if inf and ev:
            print("valid informational page")

            second_table_identifier = page.search_for(
                "Furthermore, the consultant reviewed the DNS records provided by the domains and subdomains discovered to attempt identifying if any valuable information could be obtained."
            )

            image_zoomed = pymupdf.Matrix(2.0, 2.0)
            left_point = ev[0].x0
            right_point = page.rect.width - 36

            # First screenshot: Doppelganger Domains / Sub Domain table area
            top = ev[0].y0 + 25
            bottom = page.rect.height - 80

            if second_table_identifier:
                bottom = second_table_identifier[0].y0 - 10

            table_container = pymupdf.Rect(left_point, top, right_point, bottom)
            table_text = page.get_text("text", clip=table_container).upper()

            def headers_in_clip(label: str):
                res = []
                for box in page.search_for(label):
                    if box.y0 >= (table_container.y0 - 6) and box.y1 <= (table_container.y1 + 6):
                        res.append(box)
                return res

            def same_header_row(a_boxes, b_boxes, tolerance: float = 8.0) -> bool:
                for a in a_boxes:
                    for b in b_boxes:
                        if abs(a.y0 - b.y0) <= tolerance:
                            return True
                return False

            has_misspelling_col = "MISSPELLING TECHNIQUE" in table_text
            has_domain_name_col = "DOMAIN NAME" in table_text
            has_ip_assoc_col = "IP ADDRESS ASSOCIATED WITH DOMAIN" in table_text
            has_country_col = "COUNTRY" in table_text
            is_doppelganger_table = has_misspelling_col or (has_domain_name_col and (has_ip_assoc_col or has_country_col))

            subdomain_headers = headers_in_clip("SUB DOMAIN")
            record_type_headers = headers_in_clip("RECORD TYPE")
            ip_address_headers = headers_in_clip("IP ADDRESS")
            data_headers = headers_in_clip("DATA")

            has_record_type_col = len(record_type_headers) > 0
            has_ip_address_col = len(ip_address_headers) > 0
            subdomain_with_record = same_header_row(subdomain_headers, record_type_headers)
            subdomain_with_ip = same_header_row(subdomain_headers, ip_address_headers)
            subdomain_with_data = same_header_row(subdomain_headers, data_headers)

            # Unique table identifiers for this template:
            # DNS table -> RECORD TYPE, Sub Domain table -> IP ADDRESS.
            is_dns_like_table = has_record_type_col and not is_doppelganger_table
            is_subdomain_table = has_ip_address_col and not has_record_type_col and not is_doppelganger_table

            table_kind = "unknown"
            if is_doppelganger_table:
                table_kind = "doppelganger"
            elif is_subdomain_table:
                table_kind = "subdomain"
            elif is_dns_like_table:
                table_kind = "dns_like"

            if table_kind == "doppelganger":
                first_image_path = f"./Reports/Informational/Doppelganger-Domains-Table-{i}.png"
            elif table_kind == "subdomain":
                first_image_path = f"./Reports/Informational/Sub-Domain-Table-{i}.png"
            elif table_kind == "dns_like":
                first_image_path = f"./Reports/Informational/DNS-Like-Table-{i}.png"
            else:
                first_image_path = f"./Reports/Informational/Informational-Table-{i}.png"
            pixels = page.get_pixmap(matrix=image_zoomed, clip=table_container, alpha=False)
            pixels.save(first_image_path)
            #print(f"  Page {i}: Captured '{table_kind}' table -> {first_image_path}")
            if table_kind == "subdomain":
                with Image.open(first_image_path) as img:
                    width, height = img.size
                    crop_top = int(height * 0.10)
                    if crop_top > 0 and crop_top < height:
                        img.crop((0, crop_top, width, height)).save(first_image_path)

            # Count rows in doppelganger table
            if table_kind == "doppelganger" and doppelganger_count is None:
                tables = page.find_tables(clip=table_container)
                if tables and tables.tables:
                    table = tables.tables[0]
                    # Count rows excluding header (subtract 1)
                    row_count = len(table.extract()) - 1
                    if row_count > 0:
                        doppelganger_count = row_count
                        doppelganger_page_index = i
                        #print(f"    📊 Found {doppelganger_count} doppelganger domain(s)")

            first_table_candidates.append({"page": i, "path": first_image_path, "kind": table_kind})
            if second_table_identifier and dns_image is None:
                top = second_table_identifier[0].y1 + 30
                bottom = page.rect.height - 30
                second_table_container = pymupdf.Rect(left_point, top, right_point, bottom)
                first_dns_path = f"./Reports/Informational/DNS-Record1.png"
                dns_part1, dns_part2 = capture_dns_continuation_parts(pdf, i, second_table_container, first_dns_path, image_zoomed)
                if dns_image is None:
                    dns_image = (dns_part1, dns_part2)

    used_paths = set()
    for candidate in first_table_candidates:
        if candidate["kind"] == "dns_like":
            continue
        if candidate["kind"] == "doppelganger" and doppelganger_image is None:
            doppelganger_image = candidate["path"]
            used_paths.add(candidate["path"])
        elif candidate["kind"] == "subdomain" and subdomain_image is None:
            subdomain_image = candidate["path"]
            used_paths.add(candidate["path"])

    for candidate in first_table_candidates:
        if candidate["kind"] == "dns_like":
            continue
        if doppelganger_image is None:
            doppelganger_image = candidate["path"]
            used_paths.add(candidate["path"])
            continue
        if subdomain_image is None and candidate["path"] not in used_paths:
            subdomain_image = candidate["path"]
            used_paths.add(candidate["path"])
            continue
        if doppelganger_image and subdomain_image:
            break

    pdf.close()

    doc = Document(findings_report_path)

    #print(f"\n📋 Informational Image Assignment:")
    #print(f"  Doppelganger: {doppelganger_image if doppelganger_image else 'None'}")
    #print(f"  Subdomain:    {subdomain_image if subdomain_image else 'None'}")
    if dns_image:
        dns_part1, dns_part2 = dns_image
        #print(f"  DNS Part 1:   {dns_part1}")
        #print(f"  DNS Part 2:   {dns_part2 if dns_part2 else 'None'}")
    #else:
        #print(f"  DNS:          None")
    #print(f"\n🔧 Starting document insertions...")

    inserted = False
    if doppelganger_image:
        #print(f"\n[1/3] Attempting Doppelganger insertion...")
        result = insert_image_under_heading(doc, "DOPPELGANGER DOMAINS TABLE", doppelganger_image)
        inserted = (result is not None) or inserted

    if subdomain_image:
        #print(f"\n[2/3] Attempting Subdomain insertion...")
        inserted_subdomain = insert_image_under_heading(doc, "SUB DOMAIN TABLE", subdomain_image)
        if not inserted_subdomain:
            #print(f"  Retrying with plural variant...")
            inserted_subdomain = insert_image_under_heading(doc, "SUB DOMAINS TABLE", subdomain_image)
        inserted = (inserted_subdomain is not None) or inserted

    if dns_image:
        dns_part1, dns_part2 = dns_image
        #print(f"\n[3/3] Attempting DNS insertion...")
        dns_para = insert_image_under_heading(doc, "DNS RECORD(S) TABLE", dns_part1)
        if not dns_para:
            #print(f"  Retrying with variant 'DNS RECORD TABLE'...")
            dns_para = insert_image_under_heading(doc, "DNS RECORD TABLE", dns_part1)
        if not dns_para:
            #print(f"  Retrying with variant 'DNS RECORD'...")
            dns_para = insert_image_under_heading(doc, "DNS RECORD", dns_part1)

        # Insert second DNS part if it exists and first part was successfully inserted
        if dns_para and dns_part2:
            #print(f"  📎 Inserting DNS continuation (part 2)...")
            continuation_para = add_new_paragraph(dns_para)
            indent_after_insert(dns_para, continuation_para)
            # Use resize_image for proper scaling to avoid page breaks
            inserted_ok = resize_image_modified(continuation_para, dns_part2, max_w=6.3, max_h=3.5, desired_ra=3.8)
            if inserted_ok:
                continuation_para.paragraph_format.space_before = Pt(6)
                continuation_para.paragraph_format.space_after = Pt(6)
                #print(f"    ✅ DNS part 2 inserted after part 1")
            else:
                # Fallback: remove the paragraph if image insertion failed
                parent = continuation_para._element.getparent()
                if parent is not None:
                    parent.remove(continuation_para._element)
                #print(f"    ⚠️  DNS part 2 insertion failed")

        inserted = (dns_para is not None) or inserted

    # Replace NUMBER (#) with actual doppelganger count
    if doppelganger_count is not None:
        #print(f"\n[Extra] Replacing doppelganger count placeholder...")
        number_word = number_to_word(doppelganger_count)
        replacement_text = f"{number_word} ({doppelganger_count})"

        replaced = False
        for p in doc.paragraphs:
            if "NUMBER (#)" in p.text:
                p.text = p.text.replace("NUMBER (#)", replacement_text)
                set_para(p)
                #print(f"  ✅ Replaced 'NUMBER (#)' with '{replacement_text}'")
                replaced = True
                break

        if not replaced:
            print(f"❌ Could not find 'NUMBER (#)' placeholder in Findings Report")

    doc.save(findings_report_path)
    if inserted:
        print("✅ Informational screenshots were inserted under their correct reference sections")
    else:
        print("❌ Could not find the Informational reference headings in Findings Report")

def Narrative_Exploitation(findings_report_path: str, name: str) -> None:
    target = "threats to"
    place_holder = "[CUSTOMER]"
    if name == "":
        print("❌ No customer name was provided for Narrative Exploitation")
        return
    modified_name = name + "'s"
    doc = Document(findings_report_path)
    position = None
    for p in doc.paragraphs:
        para = p.text
        if target in para and place_holder in para:
            position = p
            break
    if position is None:
        print("❌ Could not find customer place holder under Narrative, Exploitation in Findings Report")
        return
    new_val = place_host_data(position, place_holder, modified_name)
    if not new_val:
        print("❌ Could not replace the customer place holder under Narrative, Exploitation in Findings Report")
        return
    doc.save(findings_report_path)
    print("✅ Narrative Exploitation was successfully populated and saved")

def discovered_threats(technical_report: str):
    container = ""
    counter_level = Counter({"Informational": 0, "Low": 0, "Medium": 0, "High": 0, "Critical": 0})
    with pdfplumber.open(technical_report) as pdf:
        for page in pdf.pages:
            data = page.extract_text() or ""
            container = container + "\n" + data
        begin = container.upper().find("DISCOVERED THREATS")
        last = container.upper().find("MITRE ATT&CK MAPPINGS")
        if begin == -1:
            raise Exception("❌ Could not find Discovered Threats in Technical Report")
        if last == -1:
            data_holder = container[begin:]
        else:
            data_holder = container[begin:last]
        for entry in data_holder.splitlines():
            cleaned = " ".join(entry.split())
            matching = re.match(r"^(.+?)\s+(Critical|High|Medium|Low|Informational)$", cleaned, re.IGNORECASE)
            if matching:
                label_severity = matching.group(2).title()
                counter_level[label_severity] += 1
        total = dict(counter_level)
        return total

def annual_year(technical_report: str) -> int:
    container = ""
    with pdfplumber.open(technical_report) as pdf:
        for page in pdf.pages:
            data = page.extract_text() or ""
            container = container + "\n" + data
    matching = re.search(r"\b(20\d{2})\b", container)
    if not matching:
        raise Exception("❌ Could not find an annual year time stamp in Technical Report")
    year = int(matching.group(1))
    return year

def unzip_file(findings_report: str) -> Path:
    directory = tempfile.mkdtemp()
    with zipfile.ZipFile(findings_report, "r") as zip_ref:
        zip_ref.extractall(directory)
    reference = Path(directory)
    return reference

def rezip_file(folder: Path, output: str) -> None:
    folder = Path(folder)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as modified_docx:
        for file in folder.rglob("*"):
            if file.is_file():
                name_holder = file.relative_to(folder)
                modified_docx.write(file, name_holder)

def retrieve_excel_from_chart(charts: Path):
    names = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
             "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    converted = str(charts)
    data_tree = etree.parse(converted)
    content = data_tree.find(".//c:externalData", names)
    if content is None:
        return None
    corr_index = content.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    combined_names = charts.name + ".rels"
    paths = charts.parent / "_rels" / combined_names
    if not paths.exists():
        return None
    converted2 = str(paths)
    tree_relatives = etree.parse(converted2)
    root_relatives  = tree_relatives.getroot()
    for relative in root_relatives:
        rel_id = relative.get("Id")
        if rel_id == corr_index:
            target = relative.get("Target")
            excel_path = charts.parent / target
            res = excel_path.resolve()
            return res
    return None

def search_charts(document_folder: Path):
    charts_dir = document_folder/ "word" / "charts"
    files = list(charts_dir.glob("chart*.xml"))
    target_text = "total findings over time"
    chart_annual = None
    chart_severity = None
    for chart in files:
        xml_data = chart.read_text(encoding="utf-8", errors="ignore").lower()
        present_severity = True
        for label in SEVERITIES_LEVELS:
            label_mod = label.lower()
            if label_mod not in xml_data:
                present_severity = False
                break
        if present_severity and target_text not in xml_data:
            chart_severity = chart
        if target_text in xml_data:
            chart_annual = chart
    return chart_severity, chart_annual

def modify_severity_chart(excel: Path, total_count) -> None:
    workbook = openpyxl.load_workbook(excel)
    worksheet = workbook.active
    severity_cols = {}
    upper_bound = worksheet.max_row + 1
    for r in range(1, upper_bound):
        label = str(worksheet.cell(row=r, column=1).value).strip()
        if label in SEVERITIES_LEVELS:
            worksheet.cell(row=r, column=2).value = total_count[label]
    workbook.save(excel)

def modify_annual_chart(excel: Path, year: int, total_count) -> None:
    workbook = openpyxl.load_workbook(excel)
    worksheet = workbook.active
    severity_cols = {}
    upper_bound = worksheet.max_column + 1
    for c in range(1, upper_bound):
        header_info = str(worksheet.cell(row=1, column=c).value).strip()
        if header_info in SEVERITIES_LEVELS:
            severity_cols[header_info] = c
    if not severity_cols:
        raise Exception("❌ Could not find the severity labels in the yearly annual chart workbook in Findings Report")
    row_year = None
    upper_bound = worksheet.max_row + 1
    for r in range(2, upper_bound):
        data = worksheet.cell(row=r, column=1).value
        converted_data = str(data).strip()
        year_modified = str(year)
        if converted_data == year_modified:
            row_year = r
            break
    if row_year is None:
        row_year = worksheet.max_row + 1
        worksheet.cell(row=row_year, column=1).value = year
    for entry in SEVERITIES_LEVELS:
        column = severity_cols[entry]
        worksheet.cell(row=row_year, column=column).value = total_count[entry]
    workbook.save(excel)

def modify_annual_chart_range(chart: Path, workbook_path: Path) -> None:
    marker1 = ".//c:ser"
    marker2 = ".//c:tx//c:v"
    marker3 = ".//c:cat//c:f"
    marker4 = ".//c:val//c:f"
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    worksheet = workbook.active
    maximum_row = worksheet.max_row
    name = {"c" : "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    converted = str(chart)
    holder_tree = etree.parse(converted)
    holder_root = holder_tree.getroot()
    sev_columns = {}
    upper_bound = worksheet.max_column + 1
    for column in range(1, upper_bound):
        header_info = str(worksheet.cell(row=1, column=column).value).strip()
        if header_info in SEVERITIES_LEVELS:
            sev_columns[header_info] = column
    for entries in holder_root.findall(marker1, name):
        container_name1 = entries.find(marker2, name)
        if container_name1 is None:
            continue
        label = container_name1.text
        if label not in sev_columns:
            continue
        number = sev_columns[label]
        character = get_column_letter(number)
        categorical_form = entries.find(marker3, name)
        val_form = entries.find(marker4, name)
        if categorical_form is not None:
            categorical_form.text = f"Sheet1!$A$2:$A${maximum_row}"
        if val_form is not None:
            val_form.text = f"Sheet1!${character}$2:${character}${maximum_row}"
    holder_tree.write(converted, encoding="utf-8", xml_declaration=True, standalone=False)

def refresh_saved_charts_data(findings_report: str) -> None:
   file_path = str(Path(findings_report).resolve())
   word_container = win32com.client.Dispatch("Word.Application")
   word_container.Visible = False
   doc_holder = word_container.Documents.Open(file_path)
   for choice in doc_holder.InlineShapes:
       try:
           if choice.HasChart:
               choice.Chart.ChartData.Activate()
               choice.Chart.Refresh()
               choice.Chart.ChartData.Workbook.Close()
       except Exception:
            pass
   for choice in doc_holder.Shapes:
        try:
            if choice.HasChart:
                choice.Chart.ChartData.Activate()
                choice.Chart.Refresh()
                choice.Chart.ChartData.Workbook.Close()
        except Exception:
            pass
   doc_holder.Fields.Update()
   doc_holder.Save()
   doc_holder.Close()
   word_container.Quit()
   print("✅ Charts in Findings Report was refreshed successfully")

def modify_num_total_findings(document_folder: Path, total: int) -> None:
    folder = document_folder / "word"
    for xml_file in folder.rglob("*.xml"):
        data = xml_file.read_text(encoding="utf-8", errors="ignore")
        modified_data = re.sub(r"Total\s*Findings\s*:?\s*\d+", f"Total Findings: {total}", data, flags=re.IGNORECASE)
        modified_data = re.sub(r"(<a:t>\s*Total\s*Findings:?\s*</a:t>.*?<a:t>)\d+(</a:t>)", rf"\g<1>{total}\g<2>", modified_data, flags=re.IGNORECASE | re.DOTALL)
        if modified_data != data:
            xml_file.write_text(modified_data, encoding="utf-8")

def Findings_Summary(technical_report: str, findings_report_path: str) -> None:
    findings_totals = discovered_threats(technical_report)
    year = annual_year(technical_report)
    totals = sum(findings_totals.values())
    folder_document = unzip_file(findings_report_path)
    chart_severity, chart_annual = search_charts(folder_document)
    if chart_severity is None:
        print("❌ Could not find the severity chart in Findings Report")
        return
    if chart_annual is None:
        print("❌ Could not find the annual chart in Findings Report")
        return
    workbook_severity = retrieve_excel_from_chart(chart_severity)
    workbook_annual = retrieve_excel_from_chart(chart_annual)
    if workbook_severity is None:
        print("❌ Could not find the severity chart workbook in Findings Report")
        return
    if workbook_annual is None:
        print("❌ Could not find the annual chart workbook in Findings Report")
        return
    modify_severity_chart(workbook_severity, findings_totals)
    modify_annual_chart(workbook_annual, year, findings_totals)
    modify_annual_chart_range(chart_annual, workbook_annual)
    modify_num_total_findings(folder_document, totals)
    rezip_file(folder_document, findings_report_path)
    refresh_saved_charts_data(findings_report_path)
    print("✅ Findings Summary with existing charts was successfully populated and saved")

def customer_name()-> str:
    name = input("Customer Name: ").strip()
    return name

def customerName(findings_report: str, name: str, year: str) -> None:
    MARKER = False
    if name is None:
        print("❌ No customer name was provided")
        return
    if year is None:
        print("❌ No year was provided")
        return
    doc = Document(findings_report)
    for i, paragraph in enumerate(doc.paragraphs):
        if paragraph.text == "Penetration Testing":
            MARKER = True
            name_slot = doc.paragraphs[i+1]
            name_slot.clear()
            name_run = name_slot.add_run(name + "\n")
            year_run = name_slot.add_run(year)
            name_run.font.name = "Corbel"
            name_run.font.size = Pt(16)
            name_run.font.color.rgb = RGBColor(242, 101, 34)
            year_run.font.name = "Corbel"
            year_run.font.size = Pt(16)
            year_run.font.color.rgb = RGBColor(242, 101, 34)
    if MARKER:
        print("✅ Customer Name and year in title page was successfully populated and saved")
    else:
        print("❌ Could not populate Customer Name and year in title page")
    doc.save(findings_report)

def main() -> None:
    print("Automation in process...")
if __name__ == "__main__":
    main()