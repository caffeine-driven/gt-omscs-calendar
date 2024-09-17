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
    
    # Check if Time is null; if so, create an all-day event
    if pd.isna(row['Time']):
        # Handle all-day event by setting only the date
        event.begin = datetime.strptime(f"{row['Date']}", "%Y-%m-%d").date()  # Only date for all-day event
        event.make_all_day()
    else:
        # If Time is not null, create an event with specific time
        event.begin = datetime.strptime(f"{row['Date']} {row['Time']}", "%Y-%m-%d %H:%M")
    
    # Handle end date and time (optional)
    if pd.notna(row['EndDate']) and pd.isna(row['EndTime']):
        # If only EndDate is present but EndTime is null, consider it an all-day event
        event.end = datetime.strptime(f"{row['EndDate']}", "%Y-%m-%d").date()
        event.make_all_day()
    elif pd.notna(row['EndDate']) and pd.notna(row['EndTime']):
        # If both EndDate and EndTime are present, set them
        event.end = datetime.strptime(f"{row['EndDate']} {row['EndTime']}", "%Y-%m-%d %H:%M")
    
    event.description = row['Body']
    event.location = row['Location']
    
    # Add event to calendar
    calendar.events.add(event)

# Save the calendar to an ICS file
output_file = 'output/academic_calendar.ics'
with open(output_file, 'w') as ics_file:
    ics_file.writelines(calendar)

print(f"ICS file created at {output_file}")