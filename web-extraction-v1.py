from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
import os


url = 'https://www.fire.ca.gov/incidents/2025/1/7/palisades-fire/updates'
reports = []

options = Options()
# options.add_argument("--headless") # Runs browser in the background
driver = webdriver.Chrome(options=options)

driver.get("https://www.fire.ca.gov/incidents/2025/1/7/palisades-fire/updates")

# Give the page 5 seconds to load the dynamic content
time.sleep(1) 

# Now the 'detail-page' content will be there
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
detail = soup.find('div', class_='detail-page')
updates = detail.find_all('a')
# updates.reverse()

mapping = {
    "Date:": "Report_Date",
    "Time:": "Report_Time",
    "Name": "Incident_Name",
    "Start Date/Time": "Start_Date_Time",
    "Incident Status": "Incident_Status",
    "Location": "Location",
    "Type": "Type",
    "Cause": "Cause",
    "Counties": "Counties",
    "Administration Unit": "Administration_Unit",
    "Size": "Unified_Command_Agency_Size",
    "Containment": "Containment",
    "Structures Threatened": "Structures_Threatened",
    "Structures Destroyed": "Structures_Destroyed",
    "Structures Damaged": "Structures_Damaged",
    "Civilian Injuries": "Civilian_Injuries",
    "Firefighter Injuries": "Firefighter_Injuries",
    "Civilian Fatalities": "Civilian_Fatalities",
    "Firefighter Fatalities": "Firefighter_Fatalities",
    "Situation Summary": "Situation_Summary",
    "Update:": "Operational_Update",
    # "Evacuation Zone Change Notice": "Evacuation_Zone_Change_Notice", Please be advised:
    "Initial Protective Actions": "Initial_Protective_Actions",
    "Additional Resources": "Additional_Resources",
    "Incident Demographics": "Incident_Demographics_Title",
    # "Damage Assessment Text": "Damage_Assessment_Text",
    "Disaster Assessment": "Disaster_Assessment",
    "Federal Assistance": "Federal_Assistance_Text",
    "Disaster Resource Center:": "Disaster_Resource_Center_Text",
    "Westside Location": "DRC_Westside",
    "Eastside Location": "DRC_Eastside",
    # "DRC Ventura": "DRC_Ventura",
    # "Governors Office Press Releases": "Governors_Office_Press_Releases",
    "Family Assistance Center": "Family_Assistance_Text",
    "Missing Persons Hotline:": "Missing_Persons_Hotline",
    # "Public Health Information": "Public_Health_Information",
    "Press Conferences and Operational Briefings": "Operational_Briefings",
    # "Community Meeting": "Community_Meeting",
    # "Community Meeting Livestream": "Community_Meeting_Livestream",
    "Incident Photos and Videos": "Incident_Photos_Videos",
    "State Park Closures": "State_Park_Closures",
    "The Following Zones are under mandatory evacuation order:": "Evacuation_Order_Zones",
    "OPEN TO RESIDENTS ONLY!": "Evacuation_Order_Open_To_Residents",
    "Per the Los Angeles County Sheriff's Department:": "Curfew_Order",
    # "Evacuation Warning Zones": "Evacuation_Warning_Zones",
    # "Evacuation Warning Open To Residents": "Evacuation_Warning_Open_To_Residents",
    # "Evacuation Orders Details": "Evacuation_Orders_Details",
    "Evacuation Shelters": "Evacuation_Shelters",
    "Road Closures": "Road_Closure_Sources",
    "Missing Pets Line:": "Animal_Evacuation_Missing_Pets_Line",
    "Small Animals:": "Small_Animal_Shelters",
    "Large Animals:": "Large_Animal_Shelters",
    "Assigned Resources": "Assigned_Resources_Text",
    "Engines": "Engines",
    "Water Tenders": "Water_Tenders",
    "Helicopters": "Helicopters",
    "Dozers": "Dozers",
    "Hand Crews": "Hand_Crews",
    "Other": "Other",
    "Total Personnel": "Total_Personnel",
    "Cooperating Agencies": "Cooperating_Agencies",
    # "PDF File Name": "PDF_File_Name",
    # "PDF Page Count": "PDF_Page_Count"
}

report_index = 1
for update in updates:
    report = {}

    link = update['href']
    driver2 = webdriver.Chrome(options=options)
    driver2.get('https://www.fire.ca.gov' + link)
    time.sleep(1)
    html2 = driver2.page_source
    soup2 = BeautifulSoup(html2, 'html.parser')
    content = soup2.find('main', class_='main-content')

    header = content.find('h1').find_all('span')
    report["Report_Sequence"] = f"{report_index:03d}"
    report["Report_Title"] = header[0].text
    report["Incident_Name_Header"] = header[1].text
    
    rows = content.find_all('div', class_='row')
    header_cols = rows[1].find_all('div', class_='col-sm')

    # Remove hidden text
    for hidden in content.select(".visually-hidden"):
        hidden.decompose()

    agencies = "\n".join(s.get_text(strip=True) for s in header_cols[0].find_all('a'))
    report["Agencies_Listed_Top (social media links)"] = agencies

    information = header_cols[1].find_all('p') if len(header_cols) > 1 else []
    for info in information:
        span = info.find('span')
        a = info.find('a')
        if span and "Palisades Fire Public Information Line" in span.text:
            report["Public_Information_Line"] = a.text
        elif span and "Palisades Fire Media Line" in span.text:
            report["Media_Line"] = a.text
        elif span and "Online Fire Information" in span.text:
            report["Online_Fire_Information"] = a.text


    columns = content.find_all('dt')
    values = content.find_all('dd')
    for column in columns:
        value = values[columns.index(column)]
        if column.text in mapping:
            report[mapping[column.text]] = value.get_text(strip=True)

    for div in content.find_all('div', class_='p-3'):             
        h3 = div.find('h3')
        if h3 and h3.text in mapping:
            # this is missing some text, such as ul tags
            paragraphs = [p.get_text(strip=False) for p in div.find_all('p') if p.get_text(strip=True)]
            report[mapping[h3.text]] = "\n".join(paragraphs)
        
        h4 = div.find('h4')
        if h4 and h4.text in mapping:
            paragraphs = [p.get_text(strip=False) for p in div.find_all('p') if p.get_text(strip=True)]
            report[mapping[h4.text]] = "\n".join(paragraphs)


    # print(columns)
    # print(values)
    print(report)
    reports.append(report)
    report_index += 1
    driver2.quit()

    # TODO: remove this break after testing
    if report_index > 1:
        break



driver.quit()

# Export to spreadsheet
excel_filename = "palisades_fire_compiled.xlsx"
sheet_name = "Palisades_Updates"
df = pd.DataFrame(reports)
if os.path.exists(excel_filename):
    existing_headers = pd.read_excel(excel_filename, sheet_name=sheet_name, nrows=0).columns
    
    df_new = df.reindex(columns=existing_headers)
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df_new.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False, header=False)
        print(f"Excel file updated: {excel_filename}")
else:
    df.to_excel(excel_filename, sheet_name=sheet_name, index=False)
    print(f"Excel file created: {excel_filename}")

# TODO: After all assigned reports have been entered, conduct a final comprehensive check of the dataset.
# This should include:
# 1. Random report verification
# Select 5–10 reports randomly and compare them with the original reports to ensure that all information was transferred correctly.