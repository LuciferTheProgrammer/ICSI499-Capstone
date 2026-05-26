# ICSI499 Capstone - Findings Automation Tool
Applications are contained in the Source_Code folder.

AI-Enhanced Automation for Cybersecurity Reporting

This repository contains a Python-based desktop application created for the ICSI499 Capstone project. The project automates parts of the cybersecurity reporting process by extracting information from raw penetration testing reports and inserting the results into a formatted customer Findings Report.

The tool was designed to reduce the amount of manual copy-and-paste work required when preparing cybersecurity reports. It uses PDF parsing, Word document automation, Excel processing, screenshot extraction, image cropping, and AI-assisted summarization to help populate different sections of a customer report.

## Project Overview

The Findings Automation Tool takes multiple cybersecurity report files as input, including Activity, Executive, Technical, and Findings/Recommendations reports. It then extracts important information from those files and updates a Microsoft Word Findings Report template.

The application includes a graphical user interface built with `customtkinter`, allowing users to select input files and run each automation module separately. The backend automation code handles the report parsing, data extraction, screenshot creation, formatting, and report insertion.

## Main Features

- Desktop GUI built with `customtkinter`
- Automates cybersecurity Findings Report generation
- Extracts Activity Log information from Activity Reports
- Extracts Assessment Results Summary information from Executive Reports
- Uses OpenAI API for AI-assisted summary extraction
- Extracts recommendations from an Excel findings/recommendations spreadsheet
- Extracts findings details from Technical Reports
- Captures screenshots of affected nodes and evidence sections
- Crops screenshots to remove excess whitespace
- Inserts screenshots into the Findings Report
- Updates appendix sections with open ports table screenshots
- Updates customer name and report year fields
- Updates host discovery narrative information
- Updates findings summary charts and severity counts
- Supports individual module execution through the GUI
- Uses Microsoft Word automation through `win32com.client`

## Project Files

- `AutomationPrototype.py` - main backend automation code for extracting report data and inserting it into the Findings Report
- `frontend.py` - graphical user interface for selecting files and running automation modules
- `clean_crop.py` - helper script used to crop whitespace from screenshots and table images
- `Reports/` - folder containing sample input reports, templates, screenshots, and testing data

## How It Works

The application starts through `frontend.py`, which opens a desktop GUI. The user selects the required input files, including the Activity Report, Technical Report, Executive Report, Findings/Recommendations spreadsheet, and the customer Findings Report.

Each button in the GUI runs a different automation module. These modules call backend functions from `AutomationPrototype.py`.

The automation code performs tasks such as:

- Finding the Activity Log table in the Activity Report
- Extracting activity entries and inserting them into the Automated Testing Activity section
- Finding the Engagement Results Summary page in the Executive Report
- Using OpenAI to extract structured summary data from the Executive Report
- Matching findings in the Technical Report with recommendations from the spreadsheet
- Capturing affected nodes and evidence screenshots from the Technical Report
- Cropping screenshots using OpenCV-based image processing
- Inserting generated content and screenshots into the Word Findings Report
- Updating report placeholders, severity sections, appendix sections, and findings summary data

The GUI also includes an output log so users can see whether each module completed successfully or encountered an error.

## Technologies Used

- Python
- customtkinter
- tkinter
- python-docx
- pywin32 / win32com.client
- pdfplumber
- camelot-py
- PyMuPDF
- pandas
- openpyxl
- lxml
- Pillow
- OpenCV
- OpenAI API
- Microsoft Word automation
- Microsoft Excel file processing

## Purpose

This project was created for the ICSI499 Capstone course to automate repetitive cybersecurity reporting tasks. The goal was to make the reporting workflow faster, more consistent, and less error-prone by reducing manual extraction and formatting work.

The project demonstrates practical software engineering skills, including GUI development, document automation, PDF parsing, spreadsheet processing, image processing, API integration, and report generation.

## Execution Instructions

Instructions to run and execute the Findings Automation Tool:

1. Please have Python version 3.11 or higher installed
   Please have Microsoft Word installed, because the project uses win32com.client for Word automation
   NOTE: That this application only works in Windows native environment and not for Mac or Linux, please if you don't have Windows run this in a Windows machine or Windows VM to execute the application

2. Download/Clone the entire repository, ensure that all runnable code for this project such as "AutomationPrototype.py", "clean_crop.py", and "frontend.py" are all in the same directory which in here is under Source_Code folder
   
3. Once Python is installed please ensure all of these dependencies like required libraries and modules are installed:
  python -m pip install python-docx
  python -m pip install camelot-py==1.0.9
  python -m pip install pdfplumber
  python -m pip install openai
  python -m pip install pymupdf
  python -m pip install pandas
  python -m pip install openpyxl
  python -m pip install lxml
  python -m pip install pillow
  python -m pip install opencv-python
  python -m pip install customtkinter
  python -m pip install pywin32
  
  The OpenAI API key must be set as a Windows environment variable:
  setx OPENAI_API_KEY "YOUR_API_KEY_HERE"
  
  My personal OpenAI API Key for PROJECT use only:
  OPENAI_API_KEY = "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"
  
4. Once the environment is all set up, please ensure to navigate to the file path: ICSI499/Source_Code where you can run frontend.py or ICSI499/Source_Code/frontend.py or

  open PowerShell or Command Prompt and navigate to the Source_Code folder:
  cd path\to\ICSI499\Source_Code
  
  Then run the application with:
  python frontend.py
  
  Test data or inputs to use for this application are in the file path: ICSI499/Source_Code/Reports, where you have "FindingsDetailsAndRecommendations.xlsx", "OrbitalFire-ActivityReportDemo.pdf", "OrbitalFire-ExecutiveReportDemo.pdf", "OrbitalFire-TechnicalReportDemo.pdf"
  and then go to ICSI499/Source_Code/Reports/Sample499 to go to "FindingsReportTest.docx", the customer findings report. Make sure to select all of these files, based on your current working directory, and upload these files as inputs in the automation tool. Then you are ready to execute the            different 
  functionalities of the code. Also, please ensure the Findings Report is closed or else the code will not be able to get write access to the customer report. Once the automation process is complete, feel free to open the Findings Report to verify the populated sections of the customer report.

5. Finally, under ICSI499/Source_Code/Reports/Sample499/Customer_Template/FindingsReportTest.docx is the generic customer report template you can make a copy of and use that copy for further testing. Don't choose this file as the input for the Findings Report, it's meant to be a generic customer       template for all of OrbitalFire's customers.

FINALLY:

Link to Windows desktop installer for Findings Automation: https://livealbany-my.sharepoint.com/:u:/g/personal/rzadanowsky_albany_edu/IQCtdoj4iC4OQ5BC7pX18Kk6AaQFqnOi3BEgcesPBih9vus?e=IlezUE
JUST DOWNLOAD ON A WINDOWS MACHINE AND INSTALL, just need to set OpenAI API key, but no need to install all the separate Python modules. 
