# ICSI499-Capstone
AI-Enhanced Automation for Cybersecurity Reporting 

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
  
