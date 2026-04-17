import pdfplumber
from pathlib import Path
import pandas
import re

# Define the directory
pdf_dir = Path("reports")

mapping = {
    # --- INCIDENT FACTS ---
    "Name": "Incident_Name",
    "Start Date/Time": "Start_Date_Time",
    "Incident Status": "Incident_Status",
    "Location": "Location",
    "Type": "Type",
    "Cause": "Cause",
    "Counties": "Counties",
    "Administration Unit": "Administration_Unit",
    "Unified Command Agency(s)": "Unified_Command_Agency_Size", # See important note below
    "Containment": "Containment",
    "Structures Threatened": "Structures_Threatened",
    "Structures Destroyed": "Structures_Destroyed",
    "Structures Damaged": "Structures_Damaged",
    "Civilian Injuries": "Civilian_Injuries",
    "Firefighter Injuries": "Firefighter_Injuries",
    "Civilian Fatalities": "Civilian_Fatalities",
    "Firefighter Fatalities": "Firefighter_Fatalities",

    # --- ASSIGNED RESOURCES ---
    "Engines": "Engines",
    "Water Tenders": "Water_Tenders",
    "Helicopters": "Helicopters",
    "Dozers": "Dozers",
    "Hand Crews": "Hand_Crews",
    "Other": "Other",
    "Total Personnel": "Total_Personnel",
    "Cooperating Agencies": "Cooperating_Agencies",

    # --- NARRATIVE / TEXT SECTIONS ---
    "Situation Summary": "Situation_Summary",
    "Operational Update": "Operational_Update",
    "Evacuation Orders": "Evacuation_Order_Zones",
    "Evacuation Warnings": "Evacuation_Warning_Zones",
    "Evacuation Shelters": "Evacuation_Shelters",
    "Road Closures": "Road_Closure_Sources"
}

index = 0
for file in pdf_dir.glob("*.pdf"):
    report = {}

    with pdfplumber.open(file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        lines = text.split("\n")

        line_index = 0
        for line in lines:
            if line_index == 0:
                report["Report_Sequence"] = index + 1
                report["Report_Title"] = line
            elif line_index == 1:
                report["Incident_Name_Header"] = line
            elif line_index == 2:
                date = re.search("Date: .*", line)
                time = re.search("Time: .*", line)
                report["Report_Date"] = date.group(0).split(" ")[1]
                report["Report_Time"] = time.group(0).split(" ")[1]



        print(lines)
        print(report)

    index += 1

    break

        # # Strategy: Use Regex to find a pattern (e.g., "Total: $123.45")
        # # This looks for 'Total: ' followed by digits and decimals
        # match = re.search(r"Total:\s*\$(\d+\.\d{2})", text)
        
        # if match:
        #     total_amount = match.group(1)
        #     print(f"Found Total: {total_amount}")