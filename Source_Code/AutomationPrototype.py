# Automation Prototype source code.
import base64
import re
from operator import truediv
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
# from packaging.utils import NormalizedName

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
    #print("✅ All events copied.")
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
    #print("✅ All events pasted.")
    doc_holder.save(findings_report)
    print("✅ Automated Testing Activity was successfully populated and saved")

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
            #print(f"No recommendation found for {finding['finding_title']}")
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
        flame = "./Reports/gray_flame.png"
        flame_holder = body_cont1.add_run()
        flame_holder.add_picture(flame, width=Pt(12))
        body_cont1.add_run(" ")
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

# For Appendix. To insert screenshot of image after section header "SCREENSHOT OF OPEN PORTS TABLE"
def add_new_paragraph(paragraph):
    inserted = OxmlElement("w:p")
    paragraph._p.addnext(inserted)
    return Paragraph(inserted, paragraph._parent)

# To populate the Appendix section of the Findings Report.
def appendix(technical_report: str, findings_report: str) -> None:
    title = "Appendix B: Host Discovery (Opened Ports)"
    image_container = "./Reports/OPEN_PORTS.PNG"
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
    idx_screenshot = None
    position = None
    text_1 = "OPEN PORTS | EXTERNAL NETWORK TESTING"
    text_2 = "SCREENSHOT OF OPEN PORTS TABLE"
    for i, p in enumerate(doc.paragraphs):
        cur_data = p.text.upper()
        if text_1 in cur_data:
            position = i
        if text_2 in cur_data:
            idx_screenshot = i
            break
    if position is None:
        print(f"❌ {text_1} can't be found")
        return
    if idx_screenshot is None:
        print(f"❌ {text_2} can't be found")
        return
    data_container = doc.paragraphs[idx_screenshot]
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

def finding_details(technical_path: str, findings_report: str, details_path: str, sheetname: str = "External") -> None:
    dataframe = pd.read_excel(details_path, sheet_name=sheetname)
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
        find["affected_nodes_table"] = None
        find["evidence_table"] = None
        starting = find["page"]
        length_pdf = len(pdf)
        up_bound = min(starting + 4, length_pdf)
        for x in range(starting, up_bound):
            page = pdf[x]
            if find["affected_nodes_table"] is None:
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
                        file_path_affected = os.path.join(image_directory, f"Affected_Sample_{i}.png")
                        page.get_pixmap(matrix=pymupdf.Matrix(2,2), clip=crop, alpha=False).save(file_path_affected)
                        find["affected_nodes_table"] = file_path_affected
            if find["evidence_table"] is None:
                evidence = page.search_for("Evidence")
                if not evidence:
                    evidence = page.search_for("EVIDENCE")
                if evidence:
                    evidence_box = max(evidence, key=lambda s: s.y0)
                    top = evidence_box.y1 + 3
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
                        file_path_evidence = os.path.join(image_directory, f"Evidence_Sample_{i}.png")
                        page.get_pixmap(matrix=pymupdf.Matrix(2,2), clip=crop, alpha=False).save(file_path_evidence)
                        find["evidence_table"] = file_path_evidence
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
        if not cur_finds:
            continue
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
            run_para0.font.name = "Corbel"
            run_para0.font.size = Pt(12)
            run_para1 = para1.add_run(f"{i['title']}")
            run_para1.bold = True
            run_para1.font.name = "Corbel"
            run_para1.font.size = Pt(12)
            run_para_hyphen = para1.add_run(" - ")
            run_para_hyphen.font.name = "Corbel"
            run_para_hyphen.font.size = Pt(12)
            run_para2 = para1.add_run(i["description"])
            run_para2.font.name = "Corbel"
            run_para2.font.size = Pt(12)
            insert_after = para1
            counting += 1
            if i["affected_nodes_table"] and os.path.exists(i["affected_nodes_table"]):
                para_an = add_new_paragraph(insert_after)
                run_para_an = para_an.add_run("Affected Nodes:")
                run_para_an.font.italic = True
                run_para_an.font.name = "Corbel"
                run_para_an.font.size = Pt(12)
                insert_after = para_an
                para2 = add_new_paragraph(insert_after)
                resize_image(para2, i["affected_nodes_table"], max_w=6.3, max_h=2.8, desired_ra=3.8)
                insert_after = para2
            para3 = add_new_paragraph(insert_after)
            run_para3 = para3.add_run("Evidence:")
            run_para3.italic = True
            run_para3.font.name = "Corbel"
            run_para3.font.size = Pt(12)
            insert_after = para3
            if i["evidence_table"] and os.path.exists(i["evidence_table"]):
                para4 = add_new_paragraph(insert_after)
                resize_image(para4, i["evidence_table"], max_w=6.3, max_h=3.8, desired_ra=3.8)
                insert_after = para4
    doc.save(findings_report)
    print(f"✅ Findings Details was successfully populated and saved")

def Logistics(technical_path: str, findings_report: str) -> None:
    MITRE_r = []
    escalation = None
    with pdfplumber.open(technical_path) as pdf:
        MITRE_sec = False
        contact_sect = False
        for p in pdf.pages:
            data = p.extract_text()
            if not data:
                continue
            lines = data.splitlines()
            for entry in lines:
                split_entry = entry.split()
                clean_entry = " ".join(split_entry).strip()
                if clean_entry == "":
                    continue
                contact_flag = "Primary Point of Contact"
                if contact_flag in clean_entry:
                    contact_sect = True
                    continue
                if contact_sect and clean_entry.startswith("Name:"):
                    escalation = clean_entry.replace("Name:", "", 1).strip()
                    contact_sect = False
                MITRE_flag = "MITRE ATT&CK Mappings"
                if MITRE_flag in clean_entry:
                    MITRE_sec = True
                    continue
                if MITRE_sec and "Reputational Threat Findings" in clean_entry:
                    MITRE_sec = False
                if MITRE_sec:
                    if clean_entry.startswith("Time Name Tactic TTPID"):
                        continue
                    if re.search(r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun),\s+'
                                r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+'
                                r'\d{1,2},\s+\d{4}\s+@', clean_entry):
                        MITRE_r.append(clean_entry)
        if not MITRE_r:
            print("Could not find the correct rows on MITRE ATT&CK Mappings from Technical Report")
            return
        if not escalation:
            print("Could not find the correct contact for point of escalation from Technical Report")
            return
        first = MITRE_r[0]
        last = MITRE_r[-1]
        match_first = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}', first)
        match_last = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}', last)
        if not match_first or not match_last:
            print("Could not extract start and end dates from Technical Report")
            return
        start_abv = match_first.group(0)
        end_abv = match_last.group(0)
        mapping = {"Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April", "May": "May", "Jun": "June", "Jul": "July",
                   "Aug": "August", "Sep": "September", "Oct": "October", "Nov": "November", "Dec": "December"}
        starting = start_abv.replace(",", "").split()
        ending = end_abv.replace(",", "").split()
        date_start = f"{mapping[starting[0]]} {starting[1]}, {starting[2]}"
        date_end = f"{mapping[ending[0]]} {ending[1]}, {ending[2]}"
        doc = Document(findings_report)
        logistic_index = None
        struc = "The Penetration Testing was conducted according to the following:"
        for i, paragraph in enumerate(doc.paragraphs):
            if struc in paragraph.text:
                logistic_index = i
                break
        if logistic_index is None:
            print("Could not find correct Logistics section to populate in Findings Report")
            return
        bullets = [(doc.paragraphs[logistic_index + 1], "Start Date: ", date_start), (doc.paragraphs[logistic_index + 2], "End Date: ", date_end), (doc.paragraphs[logistic_index + 3], "Escalation Contact: ", escalation)]
        for y, val, s in bullets:
            y.text = val
            if y.runs:
                y.runs[0].font.name = "Corbel"
                y.runs[0].font.size = Pt(12)
            para = y.add_run(s)
            para.font.name = "Corbel"
            para.font.size = Pt(12)
            para.font.color.rgb = RGBColor(242, 101, 34)
            y.paragraph_format.space_before = Pt(0)
            y.paragraph_format.space_after = Pt(0)
    doc.save(findings_report)
    print("✅ Logistics was successfully populated and saved")

def customer_name()-> str:
    name = input("Customer Name: ").strip()
    return name

def set_customer_name(findings_report_path: str) -> None:
    name = customer_name()
    if name == "":
        print("No name was provided. Please try again.")
        return
    marker = False
    doc = Document(findings_report_path)
    cus_sec = "[CUSTOMER NAME]"
    for p in doc.paragraphs:
        if cus_sec in p.text:
            modified = p.text.replace(cus_sec, name, 1)
            p.clear()
            k = p.add_run(modified)
            k.font.name = "Corbel"
            k.font.size = Pt(16)
            k.font.color.rgb = RGBColor(230, 120, 31)
            marker = True
            break
    doc.save(findings_report_path)
    if marker:
        print("✅ Customer Name was successfully updated")
    else:
        print("❌ No Customer Name was updated")

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
            print("Could not find IP Addresses in Technical Report")
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
            print("Could not find IP Address section in Findings Report")
            return
        increment = position + 1
        index = doc.paragraphs[increment]
        for IP in collector:
            para = index.insert_paragraph_before()
            para.style = "IP"
            execute = para.add_run(IP)
            execute.font.name = "Corbel"
            execute.font.size = Pt(12)
        doc.save(findings_report_path)
        print("✅ IP Address was successfully populated and saved")

def place_host_data(paragraph, temp: str, data: str) -> bool:
    for entry in paragraph.runs:
        if temp in entry.text:
            entry.text = entry.text.replace(temp, data)
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
    SYSTEMS_HOLDER = "[SYSTEMS]"
    ADDRESSES_HOLDER = "[ADDRESSES]"
    OPEN_PORTS_HOLDER = "[OPEN PORTS]"
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
        print("Could not do data extraction from Host Discovery in Technical Report")
        return
    doc = Document(findings_report_path)
    position = None
    for p in doc.paragraphs:
        para = p.text
        if NETWORKS_HOLDER in para or SYSTEMS_HOLDER in para or ADDRESSES_HOLDER in para or OPEN_PORTS_HOLDER in para:
            position = p
            break
    if position is None:
        print("Could not find Host Discovery section in Findings Report")
        return
    new_info1 = place_host_data(position, NETWORKS_HOLDER, network_ranges)
    new_info2 = place_host_data(position, SYSTEMS_HOLDER, systems)
    new_info3 = place_host_data(position, ADDRESSES_HOLDER, addresses)
    new_info4 = place_host_data(position, OPEN_PORTS_HOLDER, open_ports)
    if not all([new_info1, new_info2, new_info3, new_info4]):
        print("Could not replace one of the placeholder values under Host Discovery in Findings Report")
        return
    doc.save(findings_report_path)
    print("✅ Host Discovery was successfully populated and saved")

def Informational(technical_report: str):
    # where header = "Informational" and Evidence:
    pdf = pymupdf.open(technical_report)
    for i in range(len(pdf)):
        page = pdf[i]
        ev = page.search_for("Evidence") # evidence part
        inf = page.search_for("Informational")
        if(inf and ev):
            print("valid informational page")
            second_table_identifier = page.search_for("Furthermore, the consultant reviewed the DNS records provided by the domains and subdomains discovered to attempt identifying if any valuable information could be obtained.")
            top = ev[0].y0 + 25
            bottom = page.rect.height - 80
            if(second_table_identifier):
                print("second table")
                bottom = second_table_identifier[0].y0 - 10
            left_point = ev[0].x0
            right_point = page.rect.width - 36
            table_container = pymupdf.Rect(left_point, top, right_point, bottom)
            image_zoomed = pymupdf.Matrix(2.0, 2.0)
            pixels = page.get_pixmap(matrix=image_zoomed, clip=table_container, alpha=False)
            # save an individual catch to reports
            pixels.save(f"./Reports/Informational{i}.png")
            if(second_table_identifier):
                print("second table")
                top = second_table_identifier[0].y1 + 30
                bottom = page.rect.height - 30
                second_table_container = pymupdf.Rect(left_point, top, right_point, bottom)
                second_pixels = page.get_pixmap(matrix=image_zoomed, clip=second_table_container, alpha=False)
                second_pixels.save(f"./Reports/Informational-second-{i}.png")

    pdf.close()
    return ""

def main() -> None:
    # activity_report, findings_report = getReports()
    #automated_testing_activity(DEFAULT_ACTIVITY_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    #assessment_results(DEFAULT_EXECUTIVE_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    #recommendations(DEFAULT_RECOMMENDATIONS_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    #recommendations(DEFAULT_RECOMMENDATIONS_PATH, DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    # locate_image_technical_report(DEFAULT_TECHNICAL_REPORT_PATH)
    #ratings = severity_counter(DEFAULT_TECHNICAL_REPORT_PATH)
    #print(f"we have {len(ratings)} vulnerabilities")
    #print(ratings)
    #finding_details(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH, DEFAULT_RECOMMENDATIONS_PATH, sheetname="External")
    #Logistics(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    #set_customer_name(DEFAULT_FINDINGS_REPORT_PATH)
    #IPAddress(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    #appendix(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    # Host_Discovery(DEFAULT_TECHNICAL_REPORT_PATH, DEFAULT_FINDINGS_REPORT_PATH)
    Informational(DEFAULT_TECHNICAL_REPORT_PATH)
if __name__ == "__main__":
    main()