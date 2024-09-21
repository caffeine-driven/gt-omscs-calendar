import pandas as pd
import requests
from ics import Calendar, Event
from datetime import datetime, date, time

from jinja2 import Environment, FileSystemLoader


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


def convert_excel_to_ics(excel_file_path: str):
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
        # Parse the start date (format is "%m/%d/%Y")
        start_date = datetime.strptime(row['Date'], "%m/%d/%Y").date()

        if pd.isna(row['Time']):  # If start time is null, treat as an all-day
            start_time = None
        else:
            start_time = datetime.strptime(row['Time'], "%H:%M").time()

        # Handle end time and end date
        if pd.notna(row['EndDate']):
            end_date = datetime.strptime(row['EndDate'], "%m/%d/%Y").date()
        else:
            end_date = None

        if pd.isna(row['EndTime']):
            end_time = None
        else:
            end_time = datetime.strptime(row['EndTime'], "%H:%M").time()

        # Calculate the duration of the event in days
        duration = (end_date - start_date).days

        if end_date and duration > 40:
            # Split the event into two separate events for the start and end dates

            # Create the first event for the start date
            event_start = create_single_day_event(
                f"{row['Title']} (Start)", start_date,
                row['Body'] if pd.notna(row['Body']) else "",
                row['Location'] if pd.notna(row['Location']) else ""
            )
            calendar.events.add(event_start)

            # Create the second event for the end date
            event_end = create_single_day_event(
                f"{row['Title']} (End)",
                end_date,
                row['Body'] if pd.notna(row['Body']) else "",
                row['Location'] if pd.notna(row['Location']) else "",
            )
            calendar.events.add(event_end)
        else:
            event = Event()
            event.name = row['Title']
            # If the event is 10 days or shorter, create a single event
            event = create_multi_day_event(
                row['Title'],
                start_date, start_time, end_date, end_time,
                row['Body'] if pd.notna(row['Body']) else "",
                row['Location'] if pd.notna(row['Location']) else ""
            )
            calendar.events.add(event)

    return calendar


def create_single_day_event(title: str, start_date: date, description: str, location: str):
    event = Event()
    event.name = title
    event.begin = start_date
    event.make_all_day()
    event.description = description
    event.location = location
    return event


def create_multi_day_event(
        title: str, start_date: date, start_time: time, end_date: date, end_time: time, description: str, location: str
):
    event = Event()
    event.name = title
    if not start_time:
        event.begin = start_date
        event.make_all_day()
    else:
        event.begin = datetime.combine(start_date, start_time)

    if not end_time:
        event.end = end_date
        event.make_all_day()
    else:
        event.end = datetime.combine(end_date, end_time)
    event.description = description
    event.location = location
    return event


def write_calendar_to_ics(calendar: Calendar, output_path: str):
    with open(output_path, 'w') as ics_file:
        ics_file.writelines(calendar)

    print(f"ICS file created at {output_path}")


def write_html(parsed_terms: list):
    env = Environment(loader=FileSystemLoader('template'))

    template = env.get_template('index.html')

    rendered_html = template.render({'items': parsed_terms})
    with open('index.html', 'w') as html_file:
        html_file.writelines(rendered_html)


if __name__ == "__main__":
    # Example usage:
    target_list = (
        ('Spring 2024', '202408'),
        ('Spring 2025', '202502')
    )
    parsed_list = []
    for term_name, code in target_list:
        excel_url = f'https://func-calendarxlsx-prod-001.azurewebsites.net/api/getExcel?termCode={code}'
        local_excel_file = f'academic_calendar_{code}.xlsx'
        output_ics_file = f'output/academic_calendar_{code}.ics'

        download_excel_file(excel_url, local_excel_file)
        cal = convert_excel_to_ics(local_excel_file)
        write_calendar_to_ics(cal, output_ics_file)
        parsed_list.append({
            'path': output_ics_file,
            'name': term_name
        })
    write_html(parsed_list)
