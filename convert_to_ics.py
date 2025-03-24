import re
from io import StringIO
from typing import Optional

import pdfplumber
import pandas as pd
import requests
from ics import Calendar, Event
from datetime import datetime, date, time

from jinja2 import Environment, FileSystemLoader


def download_file(target_url: str, save_path: str):
    """
    Downloads an Excel file from a given URL and saves it to the specified path.

    Parameters:
    target_url (str): The URL of the Excel file to download.
    save_path (str): The local path where the downloaded file will be saved.
    """
    response = requests.get(target_url)

    # Save the Excel file locally
    with open(save_path, 'wb') as file:
        file.write(response.content)

    print(f"File downloaded to {save_path}")

def read_csv(file_path: str) -> pd.DataFrame:
    return pd.read_csv(file_path, sep="\t", dtype=str)

def read_pdf(file_path: str) -> pd.DataFrame:
    all_table_rows = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row[0].strip() == 'TBD':
                        continue
                    semester = row[1].strip()
                    category = row[2].strip()
                    semester_year = int(semester.split(' ')[1])
                    dates = row[0].strip().split("-")
                    start = parse_flexible_date(dates[0], semester_year)
                    if len(dates) == 1:
                        end = None
                    else:
                        end = parse_flexible_date(dates[1], semester_year)
                    body_rows = row[3].strip().split('\n')

                    title = body_rows[0]
                    if body_rows[1:]:
                        body = "\n".join(body_rows[1:])
                    else:
                        body = None
                    all_table_rows.append([start, end, semester, category, title, body])
    return pd.DataFrame(all_table_rows, columns=['Date', 'EndDate', 'Semester', 'Category', 'Title', 'Body'])


def parse_flexible_date(date_str, current_year):
    date_str = date_str.strip()
    date_str = date_str.replace('Thur', 'Thu')
    date_str = date_str.replace('Tues', 'Tue')
    date_str = date_str.replace('(', ' (')
    date_str = date_str.replace('( ', '(')
    date_str = date_str.replace(' )', ')')
    date_str = date_str.replace('  ', ' ')
    try:
        # Try parsing with year
        return pd.to_datetime(date_str, format="%B %d, %Y (%a)").date()
    except ValueError:
        # If year is missing, assume the current year
        date_with_year = f"{date_str} {current_year}"  # Append current year
        return pd.to_datetime(date_with_year, format="%B %d (%a) %Y").date()

def convert_txt_to_ics(df: pd.DataFrame):
    # Create a calendar
    calendar = Calendar()

    # Iterate through the rows in the DataFrame and create events
    for _, row in df.iterrows():
        # Parse the start date (format is "%m/%d/%Y")
        if isinstance(row['Date'], str):
            start_date = datetime.strptime(row['Date'], "%m/%d/%Y").date()
        else:
            start_date = row['Date']

        if 'Time' not in df.columns or pd.isna(row['Time']):  # If start time is null, treat as an all-day
            start_time = None
        else:
            start_time = datetime.strptime(row['Time'], "%H:%M").time()

        # Handle end time and end date
        if pd.notna(row['EndDate']):
            if isinstance(row['EndDate'], str):
                end_date = datetime.strptime(row['EndDate'], "%m/%d/%Y").date()
            else:
                end_date = row['EndDate']
        else:
            end_date = start_date

        if 'EndTime' not in df.columns or pd.isna(row['EndTime']):
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
                get_column(row, df, 'EventLocation'),
                get_column(row, df, 'Semester'),
                get_column(row, df, 'Category'),
            )
            calendar.events.add(event_start)

            # Create the second event for the end date
            event_end = create_single_day_event(
                f"{row['Title']} (End)",
                end_date,
                row['Body'] if pd.notna(row['Body']) else "",
                get_column(row, df, 'EventLocation'),
                get_column(row, df, 'Semester'),
                get_column(row, df, 'Category'),
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
                get_column(row, df, 'EventLocation'),
                get_column(row, df, 'Semester'),
                get_column(row, df, 'Category'),
            )
            calendar.events.add(event)

    return calendar

def get_column(row, df: pd.DataFrame, column: str) -> Optional[str]:
    if column in df.columns and pd.notna(row[column]):
        return row[column]
    else:
        return None


def create_single_day_event(title: str, start_date: date, description: str, location: str, semester: str = None, category: str = None):
    event = Event()
    event.name = title
    event.begin = start_date
    event.make_all_day()
    event.description = description
    event.location = location
    categories = set()
    if semester:
        categories.add(semester)
    if category:
        categories.add(category)
    if categories:
        event.categories = categories
    return event


def create_multi_day_event(
        title: str, start_date: date, start_time: time, end_date: date, end_time: time, description: str, location: str, semester: str = None, category: str = None
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
    categories = set()
    if semester:
        categories.add(semester)
    if category:
        categories.add(category)
    if categories:
        event.categories = categories
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

def write_calendar_from_df(data_frames, output_ics_file):
    data_frame = pd.concat(data_frames, ignore_index=True)
    data_frame = data_frame.drop_duplicates()
    cal = convert_txt_to_ics(data_frame)
    write_calendar_to_ics(cal, output_ics_file)


if __name__ == "__main__":
    # Example usage:
    urls = [
        'https://registrar.gatech.edu/info/current-academic-calendar',
        'https://registrar.gatech.edu/info/future-academic-calendars'
    ]
    index = 0
    full_session_data_frames = []
    early_short_data_frames = []
    late_short_data_frames = []
    maymester_data_frames = []
    for url in urls:
        response = requests.get(url)
        response_txt = response.text
        pattern = r"https://[^\s\"']+\.pdf"
        pdf_urls = re.findall(pattern, response_txt)
        if not pdf_urls:
            continue

        for pdf_url in pdf_urls:
            local_pdf_file = f'{index}.pdf'
            download_file(pdf_url, local_pdf_file)
            index += 1
            data_frame = read_pdf(local_pdf_file)
            if 'summer' in pdf_url:
                if 'full' in pdf_url:
                    full_session_data_frames.append(data_frame)
                elif 'maymester' in pdf_url:
                    maymester_data_frames.append(data_frame)
                elif 'early' in pdf_url:
                    early_short_data_frames.append(data_frame)
                elif 'last' in pdf_url:
                    late_short_data_frames.append(data_frame)
            else:
                full_session_data_frames.append(data_frame)
                early_short_data_frames.append(data_frame)
                late_short_data_frames.append(data_frame)
                maymester_data_frames.append(data_frame)

    parsed_list = []
    if full_session_data_frames:
        ics_path = 'output/academic_calendar_full.ics'
        write_calendar_from_df(full_session_data_frames, ics_path)
        parsed_list.append({
            'path': ics_path,
            'name': "With Full Summer Semester"
        })
    if early_short_data_frames:
        ics_path = 'output/academic_calendar_early_short.ics'
        write_calendar_from_df(early_short_data_frames, ics_path)
        parsed_list.append({
            'path': ics_path,
            'name': "With Early Short Summer Semester"
        })
    if late_short_data_frames:
        ics_path = 'output/academic_calendar_late_short.ics'
        write_calendar_from_df(late_short_data_frames, ics_path)
        parsed_list.append({
            'path': ics_path,
            'name': "With Late Short Summer Semester"
        })
    if maymester_data_frames:
        ics_path = 'output/academic_calendar_maymester.ics'
        write_calendar_from_df(maymester_data_frames, ics_path)
        parsed_list.append({
            'path': ics_path,
            'name': "With Maymester Summer Semester"
        })
    write_html(parsed_list)


    # target_list = (
    #     ('Spring 2025', '202502'),
    #     ('Summer 2025', '202505'),
    # )
    # parsed_list = []
    # for term_name, code in target_list:
    #     txt_url = f'https://ro-blob.azureedge.net/ro-calendar-data/public/txt/{code}.txt'
    #     excel_url = f'https://func-calendarxlsx-prod-001.azurewebsites.net/api/getExcel?termCode={code}'
    #     local_excel_file = f'academic_calendar_{code}.xlsx'
    #     local_txt_file = f'academic_calendar_{code}.txt'
    #     output_ics_file = f'output/academic_calendar_{code}.ics'
    #
    #     # download_file(excel_url, local_excel_file)
    #     download_file(txt_url, local_txt_file)
    #     cal = convert_txt_to_ics(local_txt_file)
    #     write_calendar_to_ics(cal, output_ics_file)
    #     parsed_list.append({
    #         'path': output_ics_file,
    #         'name': term_name
    #     })
    # write_html(parsed_list)
