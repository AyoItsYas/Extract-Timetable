from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing import Callable, Iterable

    from openpyxl.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.cell import Cell

    from datetime import date, time

import argparse
import datetime
import re

import openpyxl
from icalendar import Calendar, Event
from openpyxl.cell import MergedCell


def extract_aliases(worksheet: Worksheet) -> dict:
    pattern = r"\b[A-Z]+\b"
    aliases = {}

    for cell in worksheet["B"]:
        cell: Cell
        if cell.value and type(cell.value) == str:
            if re.match(pattern, cell.value):
                x = cell.offset(column=1)
                aliases[cell.value] = x.value

    return aliases


def extract_start_point(worksheet: Worksheet) -> date:
    pattern = r"\d+"
    for cell in worksheet["B"]:
        cell: Cell
        if re.match(pattern, str(cell.value)):
            cell = cell.offset(column=1)
            return cell.value.date()


def extract_data_ranges(worksheet: Worksheet) -> tuple[str, str]:
    pattern = r"\d+"
    data_ranges = []

    for cell in worksheet["B"]:
        cell: Cell
        if cell.value and type(cell.value) == int:
            if re.match(pattern, str(cell.value)):
                x = cell.offset(row=1, column=1).coordinate
                y = cell.offset(row=8, column=5).coordinate
                yield x, y

    return data_ranges


def generate_timeframe() -> time:
    while True:
        for i in range(9, 18):
            yield datetime.time(i)


def generate_dateframe(d: date) -> date:
    delta = datetime.timedelta(days=1)

    while True:
        x = d
        d += delta

        if x.weekday() != 4:
            yield x
        else:
            d += delta * 2
            yield x


def extract_data(
    worksheet: Worksheet,
    data_ranges: Iterable[Iterable[str, str]],
) -> list[list[str]]:
    def filter_function(x, y) -> tuple[tuple[Cell]]:
        return worksheet[x:y]

    data = []
    for filter_range in data_ranges:
        data.append(filter_function(*filter_range))

    for dr in data:
        for i in range(5):
            x = []
            for row in dr:
                x.append(row[i])
            yield x


def format_summary(
    format_spec: str,
    /,
    event: str,
    aliases: dict,
    *,
    specifer_key: str = "%",
) -> str:
    specifiers = {
        "SUMMARY": event,
    }

    if ("ALIAS" in format_spec) or ("DESCRIPTION" in format_spec):
        for word in event.split():
            if word in aliases.keys():
                specifiers["SUMMARY"] = event.replace(word, aliases[word])
                specifiers["ALIAS"] = word
                specifiers["DESCRIPTION"] = aliases[word]

    for specifier, value in specifiers.items():
        format_spec = format_spec.replace(specifer_key + specifier, value)

    return format_spec


def process_calendar_events(
    data: Iterable[Iterable[Cell]],
    *,
    timeframe: Iterable[time],
    dateframe: Iterable[date],
    summary_formatter: Callable[[str], str],
) -> list[dict]:
    calendar_events = []
    datetime_constructor = lambda x, y: datetime.datetime.combine(
        x,
        y,
        tzinfo=datetime.timezone(datetime.timedelta(hours=5, minutes=30)),
    )

    for frame_date, column in zip(dateframe, data):
        for frame_time, cell in zip(timeframe, column):
            merge_flag = type(cell) is MergedCell

            start_time = frame_time
            end_time = datetime.time(start_time.hour + 1)

            if merge_flag:
                calendar_events[-1]["dtend"] = datetime_constructor(
                    frame_date, end_time
                )

            if cell.value:
                event = {
                    "summary": summary_formatter(cell.value),
                    "dtstart": datetime_constructor(frame_date, start_time),
                    "dtend": datetime_constructor(frame_date, end_time),
                }
                calendar_events.append(event)

    return calendar_events


def filter_events(
    events: Iterable[dict],
    /,
    event_filter: str,
    *,
    filter_type: str = None,
) -> Iterable[dict]:
    def filter_constructor(x: dict):
        return " ".join(y if type(y) is str else str(y) for y in x.values())

    filters = {
        "contains": lambda x: event_filter in filter_constructor(x),
        "regex": lambda x: re.match(event_filter, filter_constructor(x)),
        "!startswith": lambda x: not filter_constructor(x).startswith(event_filter),
    }

    filter_function = filters.get(filter_type)

    return filter(filter_function, events) if filter_function else events


def main(
    input_file: str,
    *,
    output_file: str,
    output_folder: str = None,
    event_filter: str,
    event_filter_type: str,
    event_format_spec: str,
) -> int:
    workbook: Workbook = openpyxl.load_workbook(input_file, data_only=True)
    worksheet: Worksheet = workbook.active

    summary = worksheet["B3"].value
    aliases = extract_aliases(worksheet)
    data_ranges = extract_data_ranges(worksheet)
    timeframe_start = extract_start_point(worksheet)

    timeframe = generate_timeframe()
    dateframe = generate_dateframe(timeframe_start)

    data = extract_data(worksheet, data_ranges=data_ranges)

    event_formatter = lambda x: format_summary(event_format_spec, x, aliases)

    calendar = Calendar()
    calendar.add("summary", summary)

    calendar_events = process_calendar_events(
        data,
        timeframe=timeframe,
        dateframe=dateframe,
        summary_formatter=event_formatter,
    )
    calendar_events = filter_events(calendar_events, "%", filter_type="!startswith")

    if event_filter:
        calendar_events = filter_events(
            calendar_events,
            event_filter,
            filter_type=event_filter_type,
        )

    for event_data in calendar_events:
        text = (
            f"{event_data['dtstart']} {event_data['dtend']} >>> {event_data['summary']}"
        )
        print(text)

        for key, value in event_data.items():
            event = Event()
            event.add(key, value)
            calendar.add_component(event)

    output_file = (output_folder or "") + (
        output_file if output_file else (summary + ".ics")
    )
    output_file = output_file.replace("%SUMMARY", summary)

    with open(output_file, "wb") as file:
        file.write(calendar.to_ical())

    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("input")
    parser.add_argument("-o", "--output")
    parser.add_argument("-of", "--output_folder")
    parser.add_argument("-f", "--filter", type=bool, default=False)
    parser.add_argument(
        "--filter_type", default="contains", choices=["contains", "regex"]
    )
    parser.add_argument("--event_format_spec", default="%ALIAS - %SUMMARY")

    args = parser.parse_args()

    status = main(
        args.input,
        output_file=args.output,
        output_folder=args.output_folder,
        event_filter=args.filter,
        event_filter_type=args.filter_type,
        event_format_spec=args.event_format_spec,
    )
    exit(status)
else:
    raise Exception("This file was not created to be imported")
