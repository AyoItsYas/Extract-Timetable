from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.workbook import Workbook
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.cell.read_only import ReadOnlyCell

    from datetime import date, time

import openpyxl
import datetime
from icalendar import Calendar, Event


def test_merge(row: int, column: int, sheet: Worksheet):
    cell = sheet.cell(row, column)
    for mergedCell in sheet.merged_cells.ranges:
        if cell.coordinate in mergedCell:
            return True
    return False


def filter_worksheet(worksheet: Worksheet) -> list[list[str]]:
    filter_ranges = (
        ("C9", "G16"),
        ("C18", "G25"),
        ("C27", "G34"),
        ("C36", "G43"),
        ("C45", "G52"),
        ("C54", "G61"),
        ("C63", "G70"),
        ("C72", "G79"),
        ("C81", "G88"),
        ("C90", "G97"),
        ("C99", "G106"),
        ("C108", "G115"),
        ("C117", "G124"),
        ("C126", "G133"),
        ("C135", "G142"),
    )

    def filter_function(x, y) -> tuple[tuple[ReadOnlyCell]]:
        return worksheet[x:y]

    data = []
    for filter_range in filter_ranges:
        x = filter_function(*filter_range)

        data.append(
            [[(b.value, test_merge(b.row, b.column, worksheet)) for b in a] for a in x]
        )

    for dr in data:
        for i in range(5):
            x = []
            for row in dr:
                x.append(row[i])
            yield x


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


def generate_timeframe() -> time:
    for i in range(9, 18):
        yield datetime.time(i)


def format_event_name(event_name: str) -> str:
    aliases = {
        "INTRO": "Introduction to Computer Science",
        "C": "Programming in C",
        "MATHS": "Mathematics for Computing",
        "DBMS": "Database Management Systems",
        "PD": "Personal Development",
        "ENG": "Academic Writing and Communication",
        "SD": "Introduction to Sustainability Development",
    }

    for word in event_name.split():
        if word in aliases.keys():
            return f"{word} - {event_name.replace(word, aliases[word])}"


def generate_calendar_data(data: list[list[str]]) -> list[dict]:
    dateframe = generate_dateframe(datetime.date(2023, 4, 17))
    calendar_data, span_flag = [], False

    for _date, events in zip(dateframe, data):
        for start_time, (event, merge_flag) in zip(generate_timeframe(), events):
            end_time = None if merge_flag else datetime.time(start_time.hour + 1)
            if span_flag:
                calendar_data[-1]["end_time"] = datetime.time(start_time.hour + 1)

            if event:
                event = format_event_name(event)

                calendar_data.append(
                    {
                        "name": event,
                        "date": _date,
                        "start_time": start_time,
                        "end_time": end_time,
                    }
                )

            span_flag = merge_flag and not span_flag

    return calendar_data


def main() -> int:
    workbook: Workbook = openpyxl.load_workbook("./data/file.xlsx", data_only=True)
    worksheet: Worksheet = workbook.active

    data = filter_worksheet(worksheet)
    calendar_data = generate_calendar_data(data)

    calendar = Calendar()

    for data in calendar_data:
        print(data["date"], data["start_time"], data["end_time"], ">" * 3, data["name"])

        event = Event()
        event.add("summary", data["name"])
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

    with open("calendar.ics", "wb") as file:
        file.write(calendar.to_ical())

    return 0


if __name__ == "__main__":
    exit(main())
else:
    raise Exception("This file was not created to be imported")
