# Automation Prototype source code.
import pdfplumber
import camelot

technical_report = input("Please enter the file path for Technical Report: ")
#executive_report = input("Please enter the file path for Executive Report: ")
#activity_report = input("Please enter the file path for Activity Report: ")
#findings_report = input("Please enter the file path for Findings Report: ")

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
    print("----------------------------------------------------------------------------------->")
    for event in activity_log:
        print(event)
def main():
    automated_testing_activity(technical_report)
main()