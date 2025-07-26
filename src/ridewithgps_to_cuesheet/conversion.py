from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from typing import List, Literal

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

from .logger import logger

# Excel formatting constants
DISTANCE_THRESHOLD_FOR_WIDE_COLUMN = 1000
CONTROL_ROW_HEIGHT: Literal[25] = 25
REGULAR_ROW_HEIGHT: Literal[15] = 15


@dataclass(frozen=True)
class EventDetails:
    name: str = "INSERT NAME OF RIDE"
    date: str = "Insert date of Ride"
    organizer: str = "Insert name of Ride Organizer"
    organizer_phone: str = "** ORGANIZER'S NUMBER **"
    start_location: str = "Insert Start location"
    finish_location: str | None = "Insert Finish location"


@dataclass(frozen=True)
class GenerationOptions:
    include_distance_from_last: bool = False
    hide_direction: bool = False
    two_decimals_precision: bool = True
    verbose: bool = False
    control_cue_indicators: List[str] = field(default_factory=lambda: ["Control", "Start", "End", "Summit"])
    end_indicator: str = "Summit"
    start_text: str = "DÉPART"
    end_text: str = "ARRIVÉE"
    page_break_row_interval: int = 40
    event_details: EventDetails = field(default_factory=EventDetails)


@dataclass(frozen=True)
class _Formats:
    title_format: Format
    description_format: Format
    control_format: Format
    arial_12: Format
    arial_12_no_border: Format
    dist_format: Format
    dist_since_format: Format
    cue_format: Format
    red_title: Format
    black_title: Format
    danger_format: Format


@dataclass(frozen=True)
class Cue:
    turn: Literal["CO", "L", "BL", "R", "BR", "TA", ""] | str
    description: str
    dist: Decimal
    is_control: bool = False
    is_danger: bool = False
    is_end: bool = False
    last_dist: Decimal = Decimal("0.0")


def generate_excel(filename: str, csv_values: List[List[str]], opts: GenerationOptions):
    cues = _parse_to_cues(csv_values, opts)
    assert cues, "No turns found in the provided CSV data."
    try:
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()

        formats = _create_excel_formats(workbook, opts.two_decimals_precision)
        last_col_letter, last_header_row = _setup_worksheet_headers(worksheet, formats, opts, cues)

        # Data processing variables
        row_num = last_header_row
        ctrl_sum = Decimal("0.0")
        last_dist = Decimal("0.0")
        page_break_list = []
        last_row_was_control = False

        for cue_num, turn in enumerate(cues):
            curr_dist = turn.dist - last_dist
            last_dist = Decimal("0.0")

            if opts.verbose:
                logger.debug(f"{turn.description}: We're on turn {cue_num} at {turn.dist}km\n"
                             f"\testimated distance is {curr_dist}km since last")

            _write_data_row(
                worksheet,
                turn,
                cue_num,
                row_num,
                last_col_letter,
                last_row_was_control,
                ctrl_sum,
                curr_dist,
                formats,
                opts,
            )

            if turn.is_control:
                ctrl_sum = Decimal("0.0")
                last_row_was_control = True
                last_dist -= curr_dist
                page_break_list.append(row_num + 1)
            else:
                last_row_was_control = False
                ctrl_sum += curr_dist
                if page_break_list and (row_num - page_break_list[-1]) == opts.page_break_row_interval:
                    page_break_list.append(row_num)

            last_dist += turn.dist
            row_num += 1

        final_row = _add_footer_information(worksheet, row_num, last_col_letter, formats)

        # Printing setup
        worksheet.print_area("A1:{0}{1}".format(last_col_letter, final_row))
        if len(page_break_list) > 2:
            page_break_list = page_break_list[1:-1]
            worksheet.set_h_pagebreaks(page_break_list)

    finally:
        if workbook is not None:
            workbook.close()


def _create_excel_formats(workbook: xlsxwriter.Workbook, two_decimals_for_dist: bool) -> _Formats:
    defaults = {"font_size": 8, "font_name": "Arial"}
    a_12_opts = {"font_size": 12, "font_name": "Arial"}
    centered = {"align": "center", "valign": "vcenter", "text_wrap": True}
    float_top = {"valign": "top"}
    all_border = {"border": 1}

    return _Formats(
        title_format=workbook.add_format({**{"rotation": 90}, **defaults, **all_border}),
        description_format=workbook.add_format({**centered, **defaults, **all_border}),
        control_format=workbook.add_format(
            {
                **{"bold": True, "bg_color": "#C0C0C0", "text_wrap": True},
                **centered,
                **a_12_opts,
                **all_border,
            }
        ),
        arial_12=workbook.add_format({**a_12_opts, **float_top, **all_border}),
        arial_12_no_border=workbook.add_format(
            {**a_12_opts, **all_border, **{"left_color": "white", "right_color": "white"}}
        ),
        dist_format=workbook.add_format(
            {**{"num_format": "0.00" if two_decimals_for_dist else "0.0"}, **float_top, **a_12_opts, **all_border}
        ),
        dist_since_format=workbook.add_format({**{"num_format": "0.0"}, **float_top, **a_12_opts, **all_border}),
        cue_format=workbook.add_format({**{"text_wrap": True}, **float_top, **a_12_opts, **all_border}),
        red_title=workbook.add_format({**{"font_color": "red"}, **a_12_opts, **centered}),
        black_title=workbook.add_format({**{"font_color": "black"}, **a_12_opts, **centered}),
        danger_format=workbook.add_format(
            {
                **{"bg_color": "#ffd700", "bold": True, "text_wrap": True},
                **a_12_opts,
                **all_border,
            }
        ),
    )


def _setup_worksheet_headers(
    worksheet: Worksheet, formats: _Formats, opts: GenerationOptions, cues: List[Cue]
) -> tuple[str, int]:
    curr_col = 0
    num_cols = 4
    if opts.include_distance_from_last:
        num_cols += 1
    if opts.hide_direction:
        num_cols -= 1

    row_num = 1
    last_col_letter = _as_letter(num_cols)

    # Header rows
    worksheet.merge_range("A1:{0}1".format(last_col_letter), opts.event_details.name, formats.red_title)
    row_num += 1
    worksheet.merge_range("A2:{0}2".format(last_col_letter), opts.event_details.date, formats.red_title)
    row_num += 1
    worksheet.merge_range("A3:{0}3".format(last_col_letter), opts.event_details.organizer, formats.red_title)
    row_num += 1
    worksheet.merge_range("A4:{0}4".format(last_col_letter), opts.event_details.start_location, formats.red_title)
    row_num += 1
    if opts.event_details.finish_location:
        worksheet.merge_range("A5:{0}5".format(last_col_letter), opts.event_details.finish_location, formats.red_title)
        row_num += 1

    # Column headers
    worksheet.write(f"A{row_num}", "Dist.(cum.)", formats.title_format)
    curr_col += 1

    if opts.include_distance_from_last:
        worksheet.write(_as_letter(curr_col) + str(row_num), "Dist. Since", formats.title_format)
        curr_col += 1

    worksheet.write(_as_letter(curr_col) + str(row_num), "Turn", formats.title_format)
    curr_col += 1

    if not opts.hide_direction:
        worksheet.write(_as_letter(curr_col) + str(row_num), "Direction", formats.title_format)
        curr_col += 1

    # Column widths
    width = 7.5 if cues[len(cues) - 1].dist > DISTANCE_THRESHOLD_FOR_WIDE_COLUMN else 6.5
    worksheet.set_column("A:A", width)
    worksheet.set_column("B:" + _as_letter(curr_col), 5.6)
    worksheet.write(_as_letter(curr_col) + str(row_num), "Route Description", formats.description_format)
    worksheet.set_column("{0}:{0}".format(_as_letter(curr_col)), 39)
    curr_col += 1

    worksheet.write(_as_letter(curr_col) + str(row_num), "Dist.(int.)", formats.title_format)
    worksheet.set_column("{0}:{0}".format(_as_letter(curr_col)), 5.6)

    return _as_letter(curr_col), row_num


def _write_data_row(
    worksheet: Worksheet,
    cue: Cue,
    cue_num: int,
    row_num: int,
    last_col_letter: str,
    last_was_control: bool,
    ctrl_sum: Decimal,
    curr_dist: Decimal,
    formats: _Formats,
    opts: GenerationOptions,
) -> None:
    curr_col = 0

    if cue_num == 1:  # no distance yet
        worksheet.write(row_num, curr_col, 0, formats.dist_format)
    else:
        prev_row = row_num
        if last_was_control and cue_num > 2:
            prev_row -= 1  # read distance from before control
        incremental_distance_formula = f"=A{prev_row}+{last_col_letter}{prev_row}"
        worksheet.write(
            row_num,
            curr_col,
            incremental_distance_formula,
            formats.dist_format,
        )
    curr_col += 1

    if opts.include_distance_from_last:
        worksheet.write(row_num, curr_col, ctrl_sum, formats.dist_since_format)
        curr_col += 1

    if cue.is_control:
        logger.info(f"Writing control at row {row_num} with description '{cue.description}'")

        worksheet.write_string(row_num, curr_col, "", formats.arial_12_no_border)
        curr_col += 1

        if not opts.hide_direction:
            worksheet.write_string(row_num, curr_col, "", formats.arial_12)
            curr_col += 1

        if cue_num == 0:  # depart, no data so far
            worksheet.merge_range(
                # n.b. convert row_num to 1-based index for Excel
                f"A{row_num + 1}:{_as_letter(curr_col - 1)}{row_num + 1}",
                "",
                formats.arial_12,
            )

        worksheet.write_string(row_num, curr_col, cue.description, formats.control_format)
        curr_col += 1
        worksheet.write_string(row_num, curr_col, "", formats.arial_12)
        worksheet.set_row(row=row_num, height=CONTROL_ROW_HEIGHT)
    else:
        worksheet.write_string(
            row_num, curr_col, cue.turn, formats.danger_format if cue.is_danger else formats.arial_12
        )
        curr_col += 1

        if not opts.hide_direction:
            worksheet.write_string(row_num, curr_col, "", formats.arial_12)
            curr_col += 1

        worksheet.write_string(
            row_num, curr_col, cue.description, formats.danger_format if cue.is_danger else formats.cue_format
        )
        curr_col += 1
        worksheet.write_number(row_num, curr_col, curr_dist, formats.dist_format)
        worksheet.set_row(row=row_num, height=REGULAR_ROW_HEIGHT)

    assert _as_letter(curr_col) == last_col_letter, "Column letter mismatch"


def _as_letter(num_after: int) -> str:
    return chr(65 + num_after)


def _add_footer_information(worksheet: Worksheet, row_num: int, last_col_letter: str, formats: _Formats) -> int:
    row_num += 1
    worksheet.merge_range(
        "A{0}:{1}{0}".format(row_num, last_col_letter),
        "IN CASE OF ABANDONMENT OR EMERGENCY",
        formats.black_title,
    )
    row_num += 1
    worksheet.merge_range(
        "A{0}:{1}{0}".format(row_num, last_col_letter),
        "PHONE: ** ORGANIZER'S NUMBER **",
        formats.black_title,
    )
    row_num += 2
    worksheet.merge_range(
        f"A{row_num}:{last_col_letter}{row_num}",
        data="TA=Turn Around, BL=Bear Left, BR=Bear Right, CO=Continue On",
        cell_format=formats.black_title,
    )
    worksheet.set_row(row=row_num - 1, height=CONTROL_ROW_HEIGHT * 2)
    return row_num + 1


def _parse_to_cues(array: List[List[str]], opts: GenerationOptions) -> List[Cue]:
    end_cue_present = False
    cues: List[Cue] = [
        None  # type: ignore # start with empty
    ] * len(array)

    last = Decimal("-1.0")
    for i in range(len(array) - 1, -1, -1):
        parsed = _read_as_cue(array[i], i, last, opts)
        cues[i] = parsed
        end_cue_present = end_cue_present or parsed.is_end
        last = parsed.last_dist

    logger.debug(f"End cue {'is' if end_cue_present else 'is not'} present in the data.")
    assert all(t for t in cues), "Some csv data could not be read as a cue."
    return cues


def _read_as_cue(row: List[str], idx: int, last_dist: Decimal, opts: GenerationOptions) -> Cue:
    has_end = False
    is_control = bool(row[0] in opts.control_cue_indicators or re.match(r"^Control.*?:", row[1]))
    is_danger = row[0].lower() == "danger"
    this_dist = Decimal(row[2])

    if idx == 1 and this_dist <= 0.1:
        this_dist = Decimal("0")

    if row[0] == opts.end_indicator:
        has_end = True
        row[1] = opts.end_text + ": " + row[1]

    return Cue(
        turn=_map_direction(row[0]),
        description=_map_cue_description(opts, row[1]).strip(),
        dist=last_dist,
        is_control=is_control,
        is_danger=is_danger,
        is_end=has_end,
        last_dist=this_dist,
    )


def _map_direction(direction: str) -> Literal["CO", "L", "BL", "R", "BR", "TA", ""] | str:
    cue_dir = direction.lower()
    if cue_dir == "straight":
        return "CO"
    elif cue_dir in ("left", "sharp left"):
        return "L"
    elif cue_dir == "slight left":
        return "BL"
    elif cue_dir in ("right", "sharp right"):
        return "R"
    elif cue_dir == "slight right":
        return "BR"
    elif cue_dir in ("generic", "food", "start", "end", "summit", "control"):
        return ""
    elif cue_dir == "uturn":
        return "TA"
    elif cue_dir == "danger":
        return "!!"
    logger.warning(f"Unknown direction '{direction}' encountered")
    return direction


def _map_cue_description(opts: GenerationOptions, description: str) -> str:
    if description == "Start of route":
        return opts.start_text
    elif description == "End of route":
        return opts.end_text
    elif match := re.match("Control.*?: *(?P<control_name>.*)", description):
        return match.group('control_name')
    elif match := re.match(r"^At roundabout, take exit (?P<exit>\d+) [io]nto (?P<road>.*)", description):
        return f"{match.group('road')} (roundabout exit {match.group('exit')})"
    elif match := re.match(r"^Continue (?:straight )?[io]nto (?P<road>.*)", description):
        return match.group('road')
    elif match := re.match(r"^(?:Keep|Turn) (?:slight )?(?:left|right) [io]nto (?P<road>.*)", description):
        return match.group('road')
    elif match := re.match(r"^(?:Make a )?U-turn on(?:to)? (?P<road>.*)", description):
        return match.group('road')
    elif match := re.match(f"Turn (?P<direction>left|right) to ([^(stay)])", description):
        return description[len(f"Turn {match.group('direction')} to ") :]

    description = description.replace("becomes", "b/c")
    description = description.replace("slightly ", "")
    return description
