#!/usr/bin/env python3

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing import Callable, Iterable

    from openpyxl.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.cell import Cell

    from datetime import time

import argparse
import datetime
import dateutil.parser
import re
from datetime import date

import openpyxl
from icalendar import Calendar, Event
from openpyxl.cell import MergedCell

NSBM_FORMAT = {
    "summary_cells": ("B3",),
    "dateframe--size": 5,
    "dateframe--start_offset": (-1, 0),
    "timeframe--range": (9, 18),
    "data_range--marker": "B",
    "data_range--marker_pattern": r"\b\d+\b",
    "data_range--point_x_offset": (1, 1),
    "data_range--point_y_offset": (8, 5),
    "alias_range--marker": "B",
    "alias_range--marker_pattern": r"\b[A-Z]+\b",
    "alias_range--offset": (0, 1),
}

PLYM_FORMAT = {
    "summary_cells": ("B3",),
    "dateframe--size": 7,
    "dateframe--start_offset": (-2, 0),
    "timeframe--range": (9, 18),
    "data_range--marker": "B",
    "data_range--marker_pattern": r"Week \d+",
    "data_range--point_x_offset": (1, 2),
    "data_range--point_y_offset": (8, 8),
    "alias_range--marker": "B",
    "alias_range--marker_pattern": r"PUSL\d{4}",
    "alias_range--offset": (0, 2),
}

PLYM_FORMAT_2 = {
    "summary_cells": ("B1", "B2"),
    "dateframe--size": 5,
    "dateframe--start_offset": (-1, 0),
    "timeframe--range": (9, 18),
    "data_range--marker": "B",
    "data_range--marker_pattern": r"Week \d+",
    "data_range--point_x_offset": (1, 2),
    "data_range--point_y_offset": (8, 8),
    "alias_range--marker": "B",
    "alias_range--marker_pattern": r"PUSL\d{4}",
    "alias_range--offset": (0, 2),
}


PLYM_FORMAT_3 = {
    "summary_cells": ("B4",),
    "dateframe--size": 5,
    "dateframe--start_offset": (-1, 0),
    "timeframe--range": (9, 18),
    "data_range--marker": "B",
    "data_range--marker_pattern": r"\d+",
    "data_range--point_x_offset": (1, 1),
    "data_range--point_y_offset": (8, 8),
    "alias_range--marker": "B",
    "alias_range--marker_pattern": r"PUSL\d{4}",
    "alias_range--offset": (0, 2),
}

NSBM_FOB_FORMAT = {
    "summary_cells": ("A2",),
    "dateframe--size": 5,
    "dateframe--start_offset": (-2, 2),
    "timeframe--range": (9, 18, 2),
    "data_range--marker": "A",
    "data_range--marker_pattern": r"Week \d+",
    "data_range--point_x_offset": (2, 2),
    "data_range--point_y_offset": (5, 6),
    "alias_range--marker": "A",
    "alias_range--marker_pattern": r"\b[A-Z]{3}_\d{4}\b",
    "alias_range--offset": (0, 1),
}

NSBM_FOB_FORMAT_VU = {
    "summary_cells": ("A1",),
    "dateframe--size": 7,
    "dateframe--start_offset": (-2, 0),
    "timeframe--range": (9, 18),
    "data_range--marker": "A",
    "data_range--marker_pattern": r"Week \d+",
    "data_range--point_x_offset": (0, 2),
    "data_range--point_y_offset": (7, 8),
    "alias_range--marker": "Z",
    "alias_range--marker_pattern": r"BM",
    "alias_range--offset": (0, 1),
}

UGC_FORMAT = {
    "summary_cells": ("B3",),
    "dateframe--size": 5,
    "dateframe--start_offset": (-1, 0),
    "timeframe--range": (9, 18),
    "data_range--marker": "B",
    "data_range--marker_pattern": r"\b\d+\b",
    "data_range--point_x_offset": (1, 1),
    "data_range--point_y_offset": (8, 5),
    "alias_range--marker": "G",
    "alias_range--marker_pattern": r"\b[A-Z]+\b",
    "alias_range--offset": (0, -3),
}

DEFINED_ANCHORS = {
    "NSBM": NSBM_FORMAT,
    "PLYM": PLYM_FORMAT,
    "PLYM_2": PLYM_FORMAT_2,
    "PLYM_3": PLYM_FORMAT_3,
    "NSBM_FOB": NSBM_FOB_FORMAT,
    "NSBM_FOB_VU": NSBM_FOB_FORMAT_VU,
    "UGC": UGC_FORMAT,
}


ANCHORS: dict = {}


def value_iterator(
    cells: Iterable[Cell],
    *,
    regex: str,
    checks: list[Callable[[Cell], bool]] = [],
) -> tuple[Cell, str]:
    for cell in cells:
        str_value = str(cell.value) if type(cell.value) is not str else cell.value
        match = re.match(regex, str_value)

        if not match:
            continue
        if any(not check(cell) for check in checks):
            continue

        match = (lambda a, b: str_value[a:b])(*match.span())
        yield cell, match


def extract_aliases(worksheet: Worksheet) -> dict:
    aliases = {}

    for cell, match in value_iterator(
        worksheet[ANCHORS["alias_range--marker"]],
        regex=ANCHORS["alias_range--marker_pattern"],
    ):
        cell: Cell
        match: str

        try:
            if len(cell.value.split()) > 1:
                continue
        except (AttributeError, ValueError, TypeError):
            continue

        cell = cell.offset(*ANCHORS["alias_range--offset"])
        aliases[match] = cell.value

    return aliases


def extract_data_ranges(worksheet: Worksheet) -> tuple[str, str]:
    for cell, match in value_iterator(
        worksheet[ANCHORS["data_range--marker"]],
        regex=ANCHORS["data_range--marker_pattern"],
    ):
        cell: Cell
        match: str

        str_value = str(cell.value) if type(cell.value) is not str else cell.value

        if match == str_value:
            x = cell.offset(*ANCHORS["data_range--point_x_offset"]).coordinate
            y = cell.offset(*ANCHORS["data_range--point_y_offset"]).coordinate
            yield x, y


def extract_dateframe_start(
    worksheet: Worksheet, data_ranges: Iterable[Iterable[str]]
) -> date:
    cords = list(data_ranges)[0][0]

    cell: Cell = worksheet[cords]

    cell = cell.offset(*ANCHORS["dateframe--start_offset"])
    try:
        return cell.value.date()
    except AttributeError:
        return dateutil.parser.parse(str(cell.value))


def generate_timeframe() -> time:
    while True:
        for i in range(*ANCHORS["timeframe--range"]):
            yield datetime.time(i)


def generate_dateframe(d: date) -> date:
    delta = datetime.timedelta(days=1)

    while True:
        x = d
        d += delta

        if ANCHORS["dateframe--size"] == 7:
            yield x
        elif x.weekday() != 4:
            yield x
        else:
            d += delta * 2
            yield x


def extract_data(
    worksheet: Worksheet,
    data_ranges: Iterable[Iterable[str, str]],
) -> list[list[str]]:
    data = []
    for filter_range in data_ranges:
        data.append((lambda x, y: worksheet[x:y])(*filter_range))

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
    specifier_key: str = "%",
) -> str:
    specifiers = {
        "SUMMARY": event,
    }

    format_spec = format_spec.split()
    format_spec = [
        x.strip(specifier_key)
        for x in filter(
            lambda x: x.startswith(specifier_key) and x.endswith(specifier_key),
            format_spec,
        )
    ]

    if ("ALIAS" in format_spec) or ("DESCRIPTION" in format_spec):
        for word in event.split():
            for alias in aliases.keys():
                if alias in word:
                    specifiers["SUMMARY"] = event.replace(word, aliases[alias])
                    specifiers["ALIAS"] = alias
                    specifiers["DESCRIPTION"] = aliases[alias]

    format_spec = [specifiers.get(spec, "") for spec in format_spec]

    return " - ".join(filter(lambda x: type(x) is str and len(x) > 0, format_spec))


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

            if type(cell.value) is str and len(cell.value) <= 2:
                continue

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
    output_folder: str,
    event_filter: str,
    event_filter_type: str,
    event_format_spec: str,
) -> int:
    workbook: Workbook = openpyxl.load_workbook(input_file, data_only=True)

    print(f"Extracting lecture schedule from workbook: {input_file}")
    n_worksheets = len(workbook.worksheets)

    for i, worksheet in enumerate(workbook.worksheets):
        print(f"Extracting from workbook {i + 1}/{n_worksheets}")

        try:
            summary_cells = [x for x in ANCHORS["summary_cells"]]
            summary = " ".join(
                str(worksheet[x].value) for x in summary_cells if worksheet[x].value
            )
            summary = summary.replace("\n", "")

            aliases = extract_aliases(worksheet)
            data_ranges = extract_data_ranges(worksheet)
            dateframe_start = extract_dateframe_start(
                worksheet, extract_data_ranges(worksheet)
            )

            timeframe = generate_timeframe()
            dateframe = generate_dateframe(dateframe_start)

            data = extract_data(worksheet, data_ranges=data_ranges)

            event_formatter = lambda event: format_summary(
                event_format_spec, event, aliases
            )

            calendar = Calendar()
            calendar.add("summary", summary)

            calendar_events = process_calendar_events(
                data,
                timeframe=timeframe,
                dateframe=dateframe,
                summary_formatter=event_formatter,
            )

            if event_filter:
                calendar_events = filter_events(
                    calendar_events,
                    event_filter,
                    filter_type=event_filter_type,
                )

            for event_data in calendar_events:
                text = f"{event_data['dtstart']} {event_data['dtend']} >>> {event_data['summary']}".replace(
                    "\n", " "
                )

                event = Event()
                for key, value in event_data.items():
                    event.add(key, value)
                calendar.add_component(event)

            def sanatize_name(path: str) -> str:
                return re.sub(r"[<>:\"/\\|*]", " ", path)

            save_path = output_folder or ""
            save_path += output_file if output_file else sanatize_name(summary + ".ics")

            save_path = save_path.replace(r"%SUMMARY%", sanatize_name(summary)[:20])
            save_path = save_path.replace(r"%WS_TITLE%", sanatize_name(worksheet.title))

            print(f"Saving to {save_path}")

            with open(save_path, "wb+") as file:
                file.write(calendar.to_ical())

            print(
                f"Successfully extracted lecture schedule '{summary} - {worksheet.title}'"
            )
        except Exception as e:
            print(
                f"Error generating iCal from worksheet: {worksheet.title} worksheet might not be a valid timetable or anchors might be mis-matched"
            )

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
    parser.add_argument("--event_format_spec", default=r"%ALIAS% %SUMMARY%")

    parser.add_argument("--anchor", default="NSBM", choices=DEFINED_ANCHORS.keys())

    args = parser.parse_args()

    ANCHORS = DEFINED_ANCHORS[args.anchor]

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
