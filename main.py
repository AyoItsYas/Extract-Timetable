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
    worksheet: Worksheet, data_ranges: Iterable[Iterable[str, str]]
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


def format_summary(event_name: str, aliases: dict) -> str:
    for word in event_name.split():
        if word in aliases.keys():
            return f"{word} - {event_name.replace(word, aliases[word])}"
    return event_name


def process_calendar_events(
    data: Iterable[Iterable[Cell]],
    *,
    timeframe: Iterable[time],
    dateframe: Iterable[date],
    summary_formatter: Callable[[str], str],
) -> list[dict]:
    calendar_events = []
    for frame_date, column in zip(dateframe, data):
        for frame_time, cell in zip(timeframe, column):
            merge_flag = type(cell) is MergedCell

            start_time = frame_time
            end_time = datetime.time(start_time.hour + 1)

            if merge_flag:
                calendar_events[-1]["end_time"] = end_time

            if cell.value:
                calendar_events.append(
                    {
                        "summary": summary_formatter(cell.value),
                        "date": frame_date,
                        "start_time": start_time,
                        "end_time": end_time,
                    }
                )

    return calendar_events


def main(
    input_file: str,
    *,
    output_file: str,
    output_folder: str = None,
    summary_filter: str,
    summary_filter_type: str,
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

    summary_formatter = lambda x: format_summary(x, aliases)

    calendar = Calendar()
    calendar.add("summary", summary or input_file)

    calendar_data = process_calendar_events(
        data,
        timeframe=timeframe,
        dateframe=dateframe,
        summary_formatter=summary_formatter,
    )

    if summary_filter:
        filter_constructor = lambda x: " ".join(
            y if type(y) is str else str(y) for y in x.values()
        )
        if summary_filter_type == "contains":
            calendar_data = filter(
                lambda x: summary_filter in filter_constructor(x), calendar_data
            )
        elif summary_filter_type == "regex":
            calendar_data = filter(
                lambda x: re.match(summary_filter, filter_constructor(x)), calendar_data
            )

    for data in calendar_data:
        text = f"{data['date']} {data['start_time']} {data['end_time']} >>> {data['summary']}"
        print(text)

        event = Event()
        event.add("summary", data["summary"])
        event.add(
            "dtstart",
            datetime.datetime.combine(
                data["date"],
                data["start_time"],
                tzinfo=datetime.timezone(datetime.timedelta(hours=5, minutes=30)),
            ),
        )
        event.add(
            "dtend",
            datetime.datetime.combine(
                data["date"],
                data["end_time"],
                tzinfo=datetime.timezone(datetime.timedelta(hours=5, minutes=30)),
            ),
        )

        calendar.add_component(event)

    output_file = (output_folder or "") + (
        output_file if output_file else (summary + ".ics")
    )
    with open(output_file, "wb") as file:
        file.write(calendar.to_ical())

    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("input")
    parser.add_argument("-o", "--output")
    parser.add_argument("-of", "--output_folder")
    parser.add_argument("-f", "--filter")
    parser.add_argument(
        "--filter_type", default="contains", choices=["contains", "regex"]
    )

    args = parser.parse_args()

    status = main(
        args.input,
        output_file=args.output,
        output_folder=args.output_folder,
        summary_filter=args.filter,
        summary_filter_type=args.filter_type,
    )
    exit(status)
else:
    raise Exception("This file was not created to be imported")
