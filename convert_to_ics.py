import pandas as pd
import requests
from ics import Calendar, Event
from datetime import datetime

# URL of the Excel file to download
url = 'https://func-calendarxlsx-prod-001.azurewebsites.net/api/getExcel?termCode=202502'

# Download the Excel file
response = requests.get(url)
excel_file_path = 'data/academic_calendar.xlsx'

# Save the Excel file locally
with open(excel_file_path, 'wb') as file:
    file.write(response.content)

# Load the downloaded Excel file into a pandas DataFrame
df = pd.read_excel(excel_file_path)

# Create a calendar
calendar = Calendar()

# Iterate through the rows in the DataFrame and create events
for _, row in df.iterrows():
    event = Event()
    event.name = row['Title']
    event.begin = datetime.strptime(f"{row['Date']} {row['Time']}", "%Y-%m-%d %H:%M")
    if pd.notna(row['EndDate']) and pd.notna(row['EndTime']):
        event.end = datetime.strptime(f"{row['EndDate']} {row['EndTime']}", "%Y-%m-%d %H:%M")
    event.description = row['Body']
    event.location = row['Location']
    calendar.events.add(event)

# Save the calendar to an ICS file
output_file = 'output/academic_calendar.ics'
with open(output_file, 'w') as ics_file:
    ics_file.writelines(calendar)

print(f"ICS file created at {output_file}")
