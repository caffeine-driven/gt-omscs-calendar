import html
import json
import re
from typing import Optional, Tuple

import pandas as pd
import requests
from ics import Calendar, Event
from datetime import datetime, date, time

from jinja2 import Environment, FileSystemLoader

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:158.0) Gecko/20100101 Firefox/158.0',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Referer': 'https://registrar.gatech.edu/current-academic-calendar',
    'X-Requested-With': 'XMLHttpRequest',
}


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


def clean_and_parse_date(raw_date_str: str, year: str) -> Tuple[date, Optional[date]]:
    # Remove <del>...</del> tags and any HTML tags
    date_str = re.sub(r'<del>.*?</del>', '', raw_date_str, flags=re.DOTALL | re.IGNORECASE)
    date_str = re.sub(r'<[^>]+>', '', date_str)
    date_str = date_str.strip()

    # Normalize full weekday names and typos to 3-letter abbreviations
    weekday_replacements = {
        'Monday': 'Mon',
        'Tuesday': 'Tue', 'Tues': 'Tue',
        'Wednesday': 'Wed',
        'Thursday': 'Thu', 'Thur': 'Thu',
        'Friday': 'Fri',
        'Saturday': 'Sat',
        'Sunday': 'Sun',
    }
    for full_name, abbr in weekday_replacements.items():
        date_str = date_str.replace(full_name, abbr)

    # Normalize parentheses spacing: e.g. "31(Fri)" -> "31 (Fri)", "( Mon)" -> "(Mon)"
    date_str = re.sub(r'\s*\(\s*', ' (', date_str)
    date_str = re.sub(r'\s*\)', ')', date_str)

    # Split date ranges by hyphens, en-dashes, or em-dashes
    dates = re.split(r'\s*[-–—]\s*', date_str)
    dates = [d.strip() for d in dates if d.strip()]

    if len(dates) > 1:
        start_month = dates[0].split()[0]
        end_parts = dates[1].split()
        if '0' <= end_parts[0][0] <= '9':
            end_date_str = f"{start_month} {dates[1]}"
        else:
            end_date_str = dates[1]
        start_date = datetime.strptime(f"{year} {dates[0]}", "%Y %B %d (%a)").date()
        end_date = datetime.strptime(f"{year} {end_date_str}", "%Y %B %d (%a)").date()
        if end_date < start_date:
            end_date = end_date.replace(year=end_date.year + 1)
    else:
        start_date = datetime.strptime(f"{year} {dates[0]}", "%Y %B %d (%a)").date()
        end_date = None

    return start_date, end_date


def clean_event_text(event_raw: str) -> str:
    # Replace line breaks or paragraph tags with spaces
    text = re.sub(r'<br\s*/?>|</p>|</div>', ' ', event_raw, flags=re.IGNORECASE)
    # Strip remaining HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    # Unescape HTML entities (&nbsp;, &quot;, &#039;, etc.)
    text = html.unescape(text)
    # Normalize multiple whitespace characters
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def get_current_academic_year() -> str:
    now = datetime.now()
    # Academic calendar year starts in May/Summer (e.g. May 2026 is 2026-2027)
    if now.month >= 5:
        default_year = f"{now.year}-{now.year + 1}"
    else:
        default_year = f"{now.year - 1}-{now.year}"

    try:
        resp = requests.get(
            'https://registrar.gatech.edu/current-academic-calendar',
            headers=HEADERS,
            timeout=15
        )
        if resp.status_code == 200:
            match = re.search(r'<select\s+id="academic-year"[^>]*>(.*?)</select>', resp.text, re.DOTALL)
            if match:
                selected_match = re.search(r'<option\s+value="([^"]+)"[^>]*selected', match.group(1))
                if selected_match:
                    return selected_match.group(1)
                all_opts = re.findall(r'<option\s+value="([^"]+)"', match.group(1))
                if all_opts:
                    return all_opts[-1]
    except Exception as e:
        print(f"Warning: Could not fetch academic year from HTML ({e}). Using default: {default_year}")

    return default_year


def fetch_calendar_data(url: str, year: str, status: str) -> list:
    try:
        resp = requests.get(url, headers=HEADERS, params={'year': year, 'status': status}, timeout=15)
        if resp.status_code == 200:
            res_json = resp.json()
            return res_json.get('data', [])
        else:
            print(f"Warning: Request for year={year}, status={status} returned HTTP {resp.status_code}")
    except Exception as e:
        print(f"Warning: Failed to fetch calendar data for year={year}, status={status}: {e}")
    return []


if __name__ == "__main__":
    url = 'https://registrar.gatech.edu/calevents/proxy'
    current_academic_year = get_current_academic_year()
    print(f"Fetching current academic year calendar: {current_academic_year}")

    # Calculate next academic year for future events
    try:
        start_y, end_y = map(int, current_academic_year.split('-'))
        future_academic_year = f"{start_y + 1}-{end_y + 1}"
    except Exception:
        future_academic_year = f"{datetime.now().year + 1}-{datetime.now().year + 2}"

    data_current = fetch_calendar_data(url, current_academic_year, 'current')
    data_future = fetch_calendar_data(url, future_academic_year, 'future')

    semesters = {
        '5A': 'Summer-All',
        '8': 'Fall',
        '5M': 'Summer-May',
        '5F': 'Summer-Full',
        '5E': 'Summer-Early',
        '5L': 'Summer-Late',
        '2': 'Spring'
    }

    def parse_event(x):
        start_date, end_date = clean_and_parse_date(x['date'], x['year'])
        event_str = clean_event_text(x['event'])
        semester_code = x.get('semester', '')
        semester_name = semesters.get(semester_code, semester_code)

        return {
            **x,
            'start_date': start_date,
            'end_date': end_date,
            'semester': semester_name,
            'event': event_str,
        }

    converted_data_current = [parse_event(item) for item in data_current]
    converted_data_future = [parse_event(item) for item in data_future]

    all_events = converted_data_current + converted_data_future
    if not all_events:
        print("Warning: No calendar events retrieved.")
        exit(0)

    df_calendar = pd.DataFrame(all_events)

    spring = df_calendar[df_calendar['semester'] == 'Spring']
    fall = df_calendar[df_calendar['semester'] == 'Fall']
    summer_full = df_calendar[df_calendar['semester'].isin(['Summer-Full', 'Summer-All'])]
    summer_early = df_calendar[df_calendar['semester'].isin(['Summer-Early', 'Summer-All'])]
    summer_late = df_calendar[df_calendar['semester'].isin(['Summer-Late', 'Summer-All'])]
    summer_maymester = df_calendar[df_calendar['semester'].isin(['Summer-May', 'Summer-All'])]

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

