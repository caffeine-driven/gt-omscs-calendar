import pandas as pd
import requests
from ics import Calendar, Event
from datetime import datetime

def download_excel_file(excel_url: str, save_path: str):
    """
    Downloads an Excel file from a given URL and saves it to the specified path.

    Parameters:
    excel_url (str): The URL of the Excel file to download.
    save_path (str): The local path where the downloaded file will be saved.
    """
    response = requests.get(excel_url)
    
    # Save the Excel file locally
    with open(save_path, 'wb') as file:
        file.write(response.content)

    print(f"Excel file downloaded to {save_path}")

def convert_excel_to_ics(excel_file_path: str, output_ics_path: str):
    """
    Converts an academic calendar from a downloaded Excel file into an ICS format.

    Parameters:
    excel_file_path (str): The local path to the downloaded Excel file.
    output_ics_path (str): The output path where the ICS file will be saved.
    """
    # Load the downloaded Excel file into a pandas DataFrame
    df = pd.read_excel(excel_file_path)

    # Create a calendar
    calendar = Calendar()

    # Iterate through the rows in the DataFrame and create events
    for _, row in df.iterrows():
        event = Event()
        event.name = row['Title']

        # Parse the start date (format is "%m/%d/%Y")
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
            
            # Calculate the duration of the event in days
            duration = (end_date - start_date).days
            
            if duration > 40:
                # Split the event into two separate events for the start and end dates

                # Create the first event for the start date
                event_start = Event()
                event_start.name = f"{row['Title']} (Start)"
                event_start.begin = start_date
                event_start.make_all_day()
                event_start.description = row['Body'] if pd.notna(row['Body']) else ""
                event_start.location = row['Location'] if pd.notna(row['Location']) else ""
                calendar.events.add(event_start)

                # Create the second event for the end date
                event_end = Event()
                event_end.name = f"{row['Title']} (End)"
                event_end.begin = end_date
                event_end.make_all_day()
                event_end.description = row['Body'] if pd.notna(row['Body']) else ""
                event_end.location = row['Location'] if pd.notna(row['Location']) else ""
                calendar.events.add(event_end)

            else:
                # If the event is 10 days or shorter, create a single event
                if pd.isna(row['EndTime']):
                    event.end = end_date
                    event.make_all_day()
                else:
                    # If end time is available, parse it
                    event.end = datetime.strptime(f"{row['EndDate']} {row['EndTime']}", "%m/%d/%Y %H:%M")

                # Add description and location
                event.description = row['Body'] if pd.notna(row['Body']) else ""
                event.location = row['Location'] if pd.notna(row['Location']) else ""

                # Add the event to the calendar
                calendar.events.add(event)

        else:
            # If no EndDate, assume it is a one-day event with no specific end time
            event.end = event.begin
            event.description = row['Body'] if pd.notna(row['Body']) else ""
            event.location = row['Location'] if pd.notna(row['Location']) else ""
            calendar.events.add(event)

    # Save the calendar to an ICS file
    with open(output_ics_path, 'w') as ics_file:
        ics_file.writelines(calendar)

    print(f"ICS file created at {output_ics_path}")

# Example usage:
excel_url = 'https://func-calendarxlsx-prod-001.azurewebsites.net/api/getExcel?termCode=202502'
local_excel_file = 'academic_calendar_202502.xlsx'
output_ics_file = 'output/academic_calendar_202502.ics'

download_excel_file(excel_url, local_excel_file)
convert_excel_to_ics(local_excel_file, output_ics_file)
