import pandas as pd
import requests
from ics import Calendar, Event
from datetime import datetime

# URL of the Excel file to download
url = 'https://func-calendarxlsx-prod-001.azurewebsites.net/api/getExcel?termCode=202502'

# Download the Excel file
response = requests.get(url)
excel_file_path = 'academic_calendar.xlsx'

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

    # Parse the date (format is "%m/%d/%Y")
    start_date = datetime.strptime(row['Date'], "%m/%d/%Y").date()
    
    if pd.isna(row['Time']):  # If start time is null, treat as an all-day event
        event.begin = start_date
        event.make_all_day()
    else:
        # If there's a specific time, parse it (assumes time is in "%H:%M" format)
        event.begin = datetime.strptime(f"{row['Date']} {row['Time']}", "%m/%d/%Y %H:%M")
    
    # Handle end time and end date
    if pd.notna(row['EndDate']):
        end_date = datetime.strptime(row['EndDate'], "%m/%d/%Y").date()
        
        if pd.isna(row['EndTime']):  # If end time is null, treat as an all-day event
            event.end = end_date
            event.make_all_day()
        else:
            # If end time is available, parse it
            event.end = datetime.strptime(f"{row['EndDate']} {row['EndTime']}", "%m/%d/%Y %H:%M")
    else:
        # If no EndDate, assume it is a one-day event with no specific end time
        event.end = event.begin

    # Add additional fields such as description and location
    event.description = row['Body'] if pd.notna(row['Body']) else ""
    event.location = row['Location'] if pd.notna(row['Location']) else ""

    # Add the event to the calendar
    calendar.events.add(event)

# Save the calendar to an ICS file
output_file = 'academic_calendar.ics'
with open(output_file, 'w') as ics_file:
    ics_file.writelines(calendar)

print(f"ICS file created at {output_file}")
