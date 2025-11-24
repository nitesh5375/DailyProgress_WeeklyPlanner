# -*- coding: utf-8 -*-
import math
import datetime
import calendar

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


def create_focus_tracker(filename="FocusTracker.xlsx"):
    # ------------------------------------------------------------------
    # CONFIG: set month and year here
    # ------------------------------------------------------------------
    year = 2025
    month = 11

    month_name = datetime.date(year, month, 1).strftime("%B")
    _, month_days = calendar.monthrange(year, month)  # correct days incl. leap years
    num_weeks = math.ceil(month_days / 7)

    # ------------------------------------------------------------------
    # Create workbook and common formats
    # ------------------------------------------------------------------
    workbook = xlsxwriter.Workbook(filename)

    title_fmt = workbook.add_format(
        {"bold": True, "font_size": 24, "align": "center", "valign": "vcenter"}
    )
    grey_header_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#EDEDED",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )
    grey_left_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#EDEDED",
            "align": "left",
            "valign": "vcenter",
            "border": 1,
        }
    )
    center_fmt = workbook.add_format({"align": "center", "valign": "vcenter"})
    border_center_fmt = workbook.add_format(
        {"align": "center", "valign": "vcenter", "border": 1}
    )
    percent_fmt = workbook.add_format({"num_format": "0%", "align": "center"})
    date_fmt = workbook.add_format(
        {"num_format": "dd.mm.yyyy", "align": "center", "valign": "vcenter"}
    )
    label_fmt = workbook.add_format({"bold": True, "align": "left"})

    # ------------------------------------------------------------------
    # Day-based color formats (shared by Habits & Week sheets)
    # ------------------------------------------------------------------
    weekday_labels = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]

    # Background colors per weekday (0=Mon..6=Sun), Sunday is red-ish
    weekday_bg_colors = {
        0: "#E3F2FD",  # Mon - light blue
        1: "#E8F5E9",  # Tue - light green
        2: "#FFF3E0",  # Wed - light orange
        3: "#F3E5F5",  # Thu - light purple
        4: "#FFFDE7",  # Fri - light yellow
        5: "#E1F5FE",  # Sat - light cyan
        6: "#FFC7CE",  # Sun - light red
    }

    # Habits formats per weekday
    habits_weekday_header_fmts = []
    habits_daynum_fmts = []
    habits_progress_fmts = []
    habits_count_fmts = []
    habits_tick_fmts = []

    # Week sheet formats per weekday
    week_date_fmts = []
    week_day_header_fmts = []
    week_tasks_header_fmts = []
    week_tick_fmts = []
    week_text_fmts = []
    week_stats_label_fmts = []
    week_stats_value_fmts = []

    for wd in range(7):
        base_bg = weekday_bg_colors[wd]

        # Sunday: stronger red header, light red data
        if wd == 6:
            header_bg = "#C00000"
            header_font = "white"
            data_bg = base_bg
            data_font = "#9C0006"
        else:
            header_bg = base_bg
            header_font = "black"
            data_bg = base_bg
            data_font = "black"

        # Habits: weekday header (MON/TUE/...), with border
        habits_weekday_header_fmts.append(
            workbook.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": header_bg,
                    "font_color": header_font,
                    "border": 1,
                }
            )
        )

        # Habits: day number row format
        habits_daynum_fmts.append(
            workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                    "border": 1,
                }
            )
        )

        # Habits: progress % row format
        habits_progress_fmts.append(
            workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                    "num_format": "0%",
                    "border": 1,
                }
            )
        )

        # Habits: Done / Not Done rows
        habits_count_fmts.append(
            workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                    "border": 1,
                }
            )
        )

        # Habits: daily habit checkbox cells
        habits_tick_fmts.append(
            workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                    "border": 1,
                }
            )
        )

        # Week sheet: date row (day of month)
        week_date_fmts.append(
            workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                }
            )
        )

        # Week sheet: day name header (MON/TUE/...), colored per day
        week_day_header_fmts.append(
            workbook.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": header_bg,
                    "font_color": header_font,
                }
            )
        )

        # Week sheet: "Tasks" header row
        week_tasks_header_fmts.append(
            workbook.add_format(
                {
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": header_bg,
                    "font_color": header_font,
                }
            )
        )

        # Week sheet: checkbox cells
        week_tick_fmts.append(
            workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                    "border": 1,
                }
            )
        )

        # Week sheet: task text cells
        week_text_fmts.append(
            workbook.add_format(
                {
                    "align": "left",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                    "border": 1,
                }
            )
        )

        # Week sheet: stats label cells (Completed / Not Completed / Total)
        week_stats_label_fmts.append(
            workbook.add_format(
                {
                    "bold": True,
                    "align": "left",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                }
            )
        )

        # Week sheet: stats value cells
        week_stats_value_fmts.append(
            workbook.add_format(
                {
                    "align": "center",
                    "valign": "vcenter",
                    "bg_color": data_bg,
                    "font_color": data_font,
                    "border": 1,
                }
            )
        )

    # ------------------------------------------------------------------
    # HABITS SHEET  (monthly habit tracker)
    # ------------------------------------------------------------------
    habits_ws = workbook.add_worksheet("Habits")

    # Column widths
    habits_ws.set_column("A:A", 2)
    habits_ws.set_column("B:B", 24)
    habits_ws.set_column("C:AG", 3)   # allow up to 31 days visually
    habits_ws.set_column("AH:AL", 10)

    # Layout parameters
    day_start_col = 2  # column C

    # Row indices (0-based)
    title_row = 1           # Excel row 2
    progress_row = 3        # Excel row 4
    done_row = 4            # Excel row 5
    not_done_row = 5        # Excel row 6
    week_header_row = 7     # Excel row 8
    weekday_row = 8         # Excel row 9
    daynum_row = 9          # Excel row 10
    habits_header_row = 10  # Excel row 11
    habit_start_row = 11    # Excel row 12

    habit_names = [
        "Start Java from 9-11PM",
        "11-1PM test already covered",
        "system design 3-4PM",
        "React 4-5PM"
    ]
    num_habits = len(habit_names)

    last_day_col = day_start_col + month_days - 1
    last_col_letter = xl_rowcol_to_cell(0, last_day_col)[:-1]

    # Title
    habits_ws.merge_range(
        title_row,
        1,
        title_row,
        last_day_col,
        f"{month_name} {year}",
        title_fmt,
    )

    # Week headers (plain grey; colors are per-day below)
    week_header_fmt = grey_header_fmt
    for w in range(num_weeks):
        start_day = w * 7 + 1
        end_day = min(month_days, start_day + 6)
        start_col = day_start_col + start_day - 1
        end_col = day_start_col + end_day - 1
        first = xl_rowcol_to_cell(week_header_row, start_col)
        last = xl_rowcol_to_cell(week_header_row, end_col)
        habits_ws.merge_range(f"{first}:{last}", f"Week {w+1}", week_header_fmt)

    # Day of week row and day numbers (colored per weekday, Sunday red)
    for day in range(1, month_days + 1):
        col = day_start_col + day - 1
        date_obj = datetime.date(year, month, day)
        wd = date_obj.weekday()  # 0..6

        habits_ws.write(
            weekday_row,
            col,
            weekday_labels[wd],
            habits_weekday_header_fmts[wd],
        )
        habits_ws.write(daynum_row, col, day, habits_daynum_fmts[wd])

    # Daily progress labels
    habits_ws.write(progress_row, 1, "Progress", grey_left_fmt)
    habits_ws.write(done_row, 1, "Done", grey_left_fmt)
    habits_ws.write(not_done_row, 1, "Not Done", grey_left_fmt)

    # "My Habits" label
    habits_ws.write(habits_header_row, 1, "My Habits", grey_header_fmt)

    # Habit list + validations (cells colored by weekday)
    for idx, name in enumerate(habit_names):
        row = habit_start_row + idx
        habits_ws.write(
            row,
            1,
            name,
            workbook.add_format(
                {
                    "bg_color": "#EDEDED",
                    "border": 1,
                    "align": "left",
                    "valign": "vcenter",
                }
            ),
        )

        for col in range(day_start_col, day_start_col + month_days):
            day_index = col - day_start_col  # 0-based
            date_obj = datetime.date(year, month, day_index + 1)
            wd = date_obj.weekday()

            habits_ws.data_validation(
                row,
                col,
                row,
                col,
                {"validate": "list", "source": ["☐", "☑"]},
            )
            habits_ws.write(row, col, "☐", habits_tick_fmts[wd])

    # Daily progress formulas (also colored by weekday)
    for offset in range(month_days):
        col = day_start_col + offset
        date_obj = datetime.date(year, month, offset + 1)
        wd = date_obj.weekday()

        top = xl_rowcol_to_cell(habit_start_row, col)
        bottom = xl_rowcol_to_cell(habit_start_row + num_habits - 1, col)

        # Done
        formula_done = f'=COUNTIF({top}:{bottom},"☑")'
        habits_ws.write_formula(done_row, col, formula_done, habits_count_fmts[wd])

        # Not done
        formula_not = f'=COUNTIF({top}:{bottom},"☐")'
        habits_ws.write_formula(not_done_row, col, formula_not, habits_count_fmts[wd])

        # % progress that day (based on filled habit cells)
        formula_p = (
            f'=IF(COUNTA({top}:{bottom})=0,"",'
            f'COUNTIF({top}:{bottom},"☑")/COUNTA({top}:{bottom}))'
        )
        habits_ws.write_formula(progress_row, col, formula_p, habits_progress_fmts[wd])

    # Analysis block on the right
    habits_ws.merge_range("AI6:AL6", "Analysis", grey_header_fmt)
    habits_ws.write("AI8", "Habit", label_fmt)
    habits_ws.write("AJ8", "Goal", label_fmt)
    habits_ws.write("AK8", "Actual", label_fmt)
    habits_ws.write("AL8", "Progress", label_fmt)

    for i, _ in enumerate(habit_names):
        analysis_row = 8 + i  # 0-based -> Excel rows 9–...
        habit_excel_row = habit_start_row + 1 + i  # linked to habit name

        # Habit name (linked)
        habits_ws.write_formula(analysis_row, 34, f"=B{habit_excel_row}")

        # Goal = number of days in the month (count of day numbers)
        goal_formula = (
            f"=COUNTA(C{daynum_row+1}:{last_col_letter}{daynum_row+1})"
        )
        habits_ws.write_formula(analysis_row, 35, goal_formula)

        # Actual completions for this habit
        top = xl_rowcol_to_cell(habit_start_row + i, day_start_col)
        bottom = xl_rowcol_to_cell(
            habit_start_row + i, day_start_col + month_days - 1
        )
        habits_ws.write_formula(
            analysis_row, 36, f'=COUNTIF({top}:{bottom},"☑")'
        )

        # Progress %
        goal_cell = xl_rowcol_to_cell(analysis_row, 35)
        actual_cell = xl_rowcol_to_cell(analysis_row, 36)
        habits_ws.write_formula(
            analysis_row,
            37,
            f'=IF({goal_cell}=0,"",{actual_cell}/{goal_cell})',
            percent_fmt,
        )

    # Charts for Habits sheet

    # Daily progress area chart (using "Done" row)
    progress_chart = workbook.add_chart({"type": "area"})
    progress_chart.add_series(
        {
            "name": "Daily Progress",
            "categories": f"=Habits!$C${daynum_row+1}:${last_col_letter}${daynum_row+1}",
            "values": f"=Habits!$C${done_row+1}:${last_col_letter}${done_row+1}",
            "fill": {"color": "#C6EFD2"},
            "border": {"none": True},
        }
    )
    progress_chart.set_legend({"none": True})
    progress_chart.set_y_axis({"major_gridlines": {"visible": False}})
    habits_ws.insert_chart("B22", progress_chart, {"x_scale": 2.5, "y_scale": 1.2})

    # Habit progress bar chart
    analysis_last_row = 8 + len(habit_names)
    bar_chart = workbook.add_chart({"type": "bar"})
    bar_chart.add_series(
        {
            "name": "Habit progress",
            "categories": f"=Habits!$AI$9:$AI${analysis_last_row}",
            "values": f"=Habits!$AL$9:$AL${analysis_last_row}",
            "fill": {"color": "#82C785"},
        }
    )
    bar_chart.set_x_axis({"num_format": "0%"})
    bar_chart.set_legend({"none": True})
    habits_ws.insert_chart("AI20", bar_chart, {"x_scale": 1.2, "y_scale": 1.4})

    # ------------------------------------------------------------------
    # WEEK SHEET  (monthly planning with daily task blocks)
    # ------------------------------------------------------------------
    week_ws = workbook.add_worksheet("Week")

    # Column widths
    week_ws.set_column("A:A", 2)
    for col in ["B", "D", "F", "H", "J", "L", "N"]:
        week_ws.set_column(f"{col}:{col}", 3)   # checkbox column
    for col in ["C", "E", "G", "I", "K", "M", "O"]:
        week_ws.set_column(f"{col}:{col}", 22)  # task text column
    week_ws.set_column("Q:R", 16)  # optional extra space

    # Title
    week_ws.merge_range("B2:O4", "Weekly Planning", title_fmt)

    # Layout for the entire month as consecutive weeks stacked vertically
    start_cols = [1, 3, 5, 7, 9, 11, 13]  # B, D, F, H, J, L, N

    tasks_per_day = 10
    # Height (rows) of one weekly block:
    #  date row (1) + day title (1) + task header (1)
    #  + tasks_per_day + completed + notcompleted + total + one blank row
    block_height = tasks_per_day + 7

    first_block_top = 8  # row index for first week's date row (Excel row 9)

    day_counter = 1

    for w in range(num_weeks):
        if day_counter > month_days:
            break

        base = first_block_top + w * block_height
        row_day_date = base
        row_day_title = base + 1
        row_task_header = base + 2
        row_task_start = base + 3
        row_completed = row_task_start + tasks_per_day
        row_notcompleted = row_completed + 1
        row_total = row_notcompleted + 1

        for i in range(7):
            if day_counter > month_days:
                break

            tick_col = start_cols[i]
            text_col = tick_col + 1

            date_obj = datetime.date(year, month, day_counter)
            wd = date_obj.weekday()  # 0..6
            day_name = date_obj.strftime("%a").upper()  # e.g. MON, TUE

            # Date row
            rng_date = (
                f"{xl_rowcol_to_cell(row_day_date, tick_col)}:"
                f"{xl_rowcol_to_cell(row_day_date, text_col)}"
            )
            week_ws.merge_range(rng_date, date_obj.day, week_date_fmts[wd])

            # Day title row
            rng_title = (
                f"{xl_rowcol_to_cell(row_day_title, tick_col)}:"
                f"{xl_rowcol_to_cell(row_day_title, text_col)}"
            )
            week_ws.merge_range(rng_title, day_name, week_day_header_fmts[wd])

            # Tasks header row
            rng_tasks = (
                f"{xl_rowcol_to_cell(row_task_header, tick_col)}:"
                f"{xl_rowcol_to_cell(row_task_header, text_col)}"
            )
            week_ws.merge_range(rng_tasks, "Tasks", week_tasks_header_fmts[wd])

            # Tasks area (checkbox + text), colored by weekday
            top_tick = row_task_start
            bottom_tick = row_task_start + tasks_per_day - 1
            for r in range(top_tick, bottom_tick + 1):
                week_ws.data_validation(
                    r,
                    tick_col,
                    r,
                    tick_col,
                    {"validate": "list", "source": ["☐", "☑"]},
                )
                week_ws.write(r, tick_col, "☐", week_tick_fmts[wd])
                week_ws.write(r, text_col, "", week_text_fmts[wd])

            tick_top_addr = xl_rowcol_to_cell(top_tick, tick_col)
            tick_bot_addr = xl_rowcol_to_cell(bottom_tick, tick_col)
            task_top_addr = xl_rowcol_to_cell(top_tick, text_col)
            task_bot_addr = xl_rowcol_to_cell(bottom_tick, text_col)

            # Stats rows (labels + values), colored by weekday
            week_ws.write(
                row_completed,
                tick_col,
                "Completed",
                week_stats_label_fmts[wd],
            )
            week_ws.write(
                row_completed,
                text_col,
                f'=COUNTIF({tick_top_addr}:{tick_bot_addr},"☑")',
                week_stats_value_fmts[wd],
            )

            week_ws.write(
                row_notcompleted,
                tick_col,
                "Not Completed",
                week_stats_label_fmts[wd],
            )
            week_ws.write(
                row_notcompleted,
                text_col,
                f'=COUNTIF({tick_top_addr}:{tick_bot_addr},"☐")',
                week_stats_value_fmts[wd],
            )

            week_ws.write(
                row_total,
                tick_col,
                "Total tasks",
                week_stats_label_fmts[wd],
            )
            week_ws.write(
                row_total,
                text_col,
                f"=COUNTA({task_top_addr}:{task_bot_addr})",
                week_stats_value_fmts[wd],
            )

            day_counter += 1

    workbook.close()


if __name__ == "__main__":
    create_focus_tracker()