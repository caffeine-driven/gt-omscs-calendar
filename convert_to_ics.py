import json
from typing import Optional

import pdfplumber
import pandas as pd
import requests
from ics import Calendar, Event
from datetime import datetime, date, time

from jinja2 import Environment, FileSystemLoader


def convert_txt_to_ics(df: pd.DataFrame):
    # Create a calendar
    calendar = Calendar()

    # Iterate through the rows in the DataFrame and create events
    for _, row in df.iterrows():
        start_date = row['start_date']
        end_date = row['end_date']
        if not end_date:
            end_date = start_date

        # Calculate the duration of the event in days
        duration = (end_date - start_date).days

        if end_date and duration > 10:
            # Split the event into two separate events for the start and end dates

            # Create the first event for the start date
            event_start = create_single_day_event(
                f"{row['event']} (Start)", start_date,
                semester=get_column(row, df, 'semester'),
                category=get_column(row, df, 'category'),
            )
            calendar.events.add(event_start)

            # Create the second event for the end date
            event_end = create_single_day_event(
                f"{row['event']} (End)",
                end_date,
                semester=get_column(row, df, 'semester'),
                category=get_column(row, df, 'category'),
            )
            calendar.events.add(event_end)
        else:
            event = Event()
            event.name = row['event']
            # If the event is 10 days or shorter, create a single event
            event = create_multi_day_event(
                row['event'],
                start_date, None, end_date, None,
                semester=get_column(row, df, 'semester'),
                category=get_column(row, df, 'category'),
            )
            calendar.events.add(event)

    return calendar

def get_column(row, df: pd.DataFrame, column: str) -> Optional[str]:
    if column in df.columns and pd.notna(row[column]):
        return row[column]
    else:
        return None


def create_single_day_event(
        title: str, start_date: date,
        description: Optional[str] = None, location: Optional[str] = None,
        semester: Optional[str] = None, category: Optional[str] = None
):
    event = Event()
    event.name = f'{title} - {semester}'
    event.begin = start_date
    event.make_all_day()
    event.description = ""
    if description:
        event.description = description
    if location:
        event.location = location
    categories = set()
    if semester:
        categories.add(semester)
    if category:
        categories.add(category)
    if categories:
        event.categories = categories
        event.description += "\n"+",".join(categories)
    return event


def create_multi_day_event(
        title: str, start_date: date, start_time: Optional[time], end_date: date, end_time: Optional[time],
        description: Optional[str] = None, location: Optional[str] = None,
        semester: Optional[str] = None, category: Optional[str] = None
):
    event = Event()
    event.name = f'{title} - {semester}'
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
    event.description = ""
    if description:
        event.description = description
    if location:
        event.location = location
    categories = set()
    if semester:
        categories.add(semester)
    if category:
        categories.add(category)
    if categories:
        event.categories = categories
        event.description += "\n"+",".join(categories)
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

    url = 'https://registrar.gatech.edu/calevents/proxy'
    current_year = datetime.now().year
    next_year = current_year + 1
    query = {
        "year":f"{current_year}-{next_year}",
        "status":"current"
    }
    response = requests.get(url, headers={
        'Referer':'https://registrar.gatech.edu/current-academic-calendar'
    }, params=query)

    res_data = response.json()
    semesters = {
        '5A': 'Summer-All',
        '8': 'Fall',
        '5M': 'Summer-May',
        '5F': 'Summer-Full',
        '5E': 'Summer-Early',
        '5L': 'Summer-Late',
        '2': 'Spring'
    }
    data = res_data['data']
    def parse_date(x):
        date_str = x['date'].strip()
        date_str = date_str.replace('Thur', 'Thu')
        date_str = date_str.replace('Tues', 'Tue')
        date_str = date_str.replace('( ', '(')
        date_str = date_str.replace(' )', ')')
        if '–' in date_str:
            dates = date_str.split(' – ')
        else:
            dates = date_str.split(' - ')
        if len(dates) > 1:
            start_month = dates[0].split()[0]
            end_date_str = dates[1].split()[0]
            if '0' <= end_date_str[0] <= '9':
                end_date_str = f'{start_month} {dates[1]}'
            else:
                end_date_str = dates[1]
            start_date = datetime.strptime(f'{x["year"]} {dates[0]}', "%Y %B %d (%a)").date()
            end_date = datetime.strptime(f'{x["year"]} {end_date_str}', "%Y %B %d (%a)").date()
        else:
            start_date = datetime.strptime(f'{x["year"]} {dates[0]}', "%Y %B %d (%a)").date()
            end_date = None
        event_str = x['event']
        event_str = event_str.replace('<p>', '')
        event_str = event_str.replace('</p>', '')

        return {
            **x,
            'start_date': start_date,
            'end_date': end_date,
            'semester': semesters[x['semester']],
            'event': event_str,
        }
    converted_data = list(map(parse_date, data))

    df_calendar = pd.DataFrame(converted_data)

    spring = df_calendar[df_calendar['semester'] == 'Spring']
    fall = df_calendar[df_calendar['semester'] == 'Fall']
    summer_full = df_calendar[df_calendar['semester'] == 'Summer-Full']
    summer_early = df_calendar[df_calendar['semester'] == 'Summer-Early']
    summer_late = df_calendar[df_calendar['semester'] == 'Summer-Late']
    summer_maymester = df_calendar[df_calendar['semester'] == 'Summer-May']

    full_session_data_frames = [spring, fall, summer_full]
    early_short_data_frames = [spring, fall, summer_early]
    late_short_data_frames = [spring, fall, summer_late]
    maymester_data_frames = [spring, fall, summer_maymester]


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
