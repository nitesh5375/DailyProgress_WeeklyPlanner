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
    # ROW-BASED COLOR PALETTES
    # ------------------------------------------------------------------

    # Habits: each HABIT ROW gets its own color (horizontal bands)
    habit_row_colors = [
        "#E3F2FD",  # 1
        "#E8F5E9",  # 2
        "#FFF3E0",  # 3
        "#F3E5F5",  # 4
        "#FFFDE7",  # 5
        "#E1F5FE",  # 6
        "#FCE4EC",  # 7
        "#E0F7FA",  # 8
        "#FFEBEE",  # 9
        "#F9FBE7",  # 10
    ]
    habit_name_fmts = []
    habit_tick_fmts = []
    for color in habit_row_colors:
        habit_name_fmts.append(
            workbook.add_format(
                {
                    "bg_color": color,
                    "border": 1,
                    "align": "left",
                    "valign": "vcenter",
                }
            )
        )
        habit_tick_fmts.append(
            workbook.add_format(
                {
                    "bg_color": color,
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
        )

    # Habits: weekday / day-number header formats, with Sunday red
    weekday_header_fmt = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )
    weekday_header_sun_fmt = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_color": "red",
        }
    )
    daynum_fmt = workbook.add_format(
        {
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )
    daynum_sun_fmt = workbook.add_format(
        {
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_color": "red",
        }
    )

    # Week sheet: each TASK ROW gets its own color (horizontal bands)
    week_row_colors = [
        "#E3F2FD",
        "#E8F5E9",
        "#FFF3E0",
        "#F3E5F5",
        "#FFFDE7",
        "#E1F5FE",
        "#FCE4EC",
        "#E0F7FA",
        "#FFEBEE",
        "#F9FBE7",
    ]
    week_tick_row_fmts = []
    week_text_row_fmts = []
    for color in week_row_colors:
        week_tick_row_fmts.append(
            workbook.add_format(
                {
                    "bg_color": color,
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
        )
        week_text_row_fmts.append(
            workbook.add_format(
                {
                    "bg_color": color,
                    "border": 1,
                    "align": "left",
                    "valign": "vcenter",
                }
            )
        )

    # Week sheet: day header / date formats (Sunday red)
    week_day_header_fmt = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#E0E0E0",
            "font_color": "black",
        }
    )
    week_day_header_sun_fmt = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#C00000",
            "font_color": "white",
        }
    )
    week_date_fmt_normal = workbook.add_format(
        {
            "align": "center",
            "valign": "vcenter",
        }
    )
    week_date_fmt_sun = workbook.add_format(
        {
            "align": "center",
            "valign": "vcenter",
            "font_color": "red",
        }
    )
    week_tasks_header_fmt = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#B0BEC5",
            "font_color": "black",
        }
    )
    week_tasks_header_sun_fmt = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#FFCDD2",
            "font_color": "black",
        }
    )

    weekday_labels = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]

    # ------------------------------------------------------------------
    # HABITS SHEET
    # ------------------------------------------------------------------
    habits_ws = workbook.add_worksheet("Habits")

    habits_ws.set_column("A:A", 2)
    habits_ws.set_column("B:B", 24)
    habits_ws.set_column("C:AG", 3)
    habits_ws.set_column("AH:AL", 10)

    # Layout parameters
    day_start_col = 2  # column C

    # Row indices (0‑based)
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
        "Wake up at 05:00 ⏰",
        "Gym 💪",
        "Reading / Learning 📚",
        "Day Planning 📅",
        "Budget Tracking 💰",
        "Project Work 🎯",
        "No Alcohol 🍾",
        "Social Media Detox 🌿",
        "Goal Journaling 📒",
        "Cold Shower 🚿",
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

    # Week headers (Week 1, Week 2, ...) – neutral grey
    for w in range(num_weeks):
        start_day = w * 7 + 1
        end_day = min(month_days, start_day + 6)
        start_col = day_start_col + start_day - 1
        end_col = day_start_col + end_day - 1
        first = xl_rowcol_to_cell(week_header_row, start_col)
        last = xl_rowcol_to_cell(week_header_row, end_col)
        habits_ws.merge_range(f"{first}:{last}", f"Week {w+1}", grey_header_fmt)

    # Weekday row + day numbers (Sunday red text)
    for day in range(1, month_days + 1):
        col = day_start_col + day - 1
        date_obj = datetime.date(year, month, day)
        wd = date_obj.weekday()  # 0..6

        if wd == 6:
            habits_ws.write(weekday_row, col, weekday_labels[wd], weekday_header_sun_fmt)
            habits_ws.write(daynum_row, col, day, daynum_sun_fmt)
        else:
            habits_ws.write(weekday_row, col, weekday_labels[wd], weekday_header_fmt)
            habits_ws.write(daynum_row, col, day, daynum_fmt)

    # Daily progress labels
    habits_ws.write(progress_row, 1, "Progress", grey_left_fmt)
    habits_ws.write(done_row, 1, "Done", grey_left_fmt)
    habits_ws.write(not_done_row, 1, "Not Done", grey_left_fmt)

    # "My Habits" label
    habits_ws.write(habits_header_row, 1, "My Habits", grey_header_fmt)

    # Habit rows (each row a different color horizontally)
    for idx, name in enumerate(habit_names):
        row = habit_start_row + idx
        color_idx = idx % len(habit_row_colors)

        # Habit name
        habits_ws.write(row, 1, name, habit_name_fmts[color_idx])

        # Habit cells per day
        for col in range(day_start_col, day_start_col + month_days):
            habits_ws.data_validation(
                row,
                col,
                row,
                col,
                {"validate": "list", "source": ["☐", "☑"]},
            )
            habits_ws.write(row, col, "☐", habit_tick_fmts[color_idx])

    # Daily progress formulas (summary rows)
    for offset in range(month_days):
        col = day_start_col + offset
        top = xl_rowcol_to_cell(habit_start_row, col)
        bottom = xl_rowcol_to_cell(habit_start_row + num_habits - 1, col)

        # Done
        formula_done = f'=COUNTIF({top}:{bottom},"☑")'
        habits_ws.write_formula(done_row, col, formula_done, border_center_fmt)

        # Not done
        formula_not = f'=COUNTIF({top}:{bottom},"☐")'
        habits_ws.write_formula(not_done_row, col, formula_not, border_center_fmt)

        # % progress that day
        formula_p = (
            f'=IF(COUNTA({top}:{bottom})=0,"",'
            f'COUNTIF({top}:{bottom},"☑")/COUNTA({top}:{bottom}))'
        )
        habits_ws.write_formula(progress_row, col, formula_p, percent_fmt)

    # Analysis block
    habits_ws.merge_range("AI6:AL6", "Analysis", grey_header_fmt)
    habits_ws.write("AI8", "Habit", label_fmt)
    habits_ws.write("AJ8", "Goal", label_fmt)
    habits_ws.write("AK8", "Actual", label_fmt)
    habits_ws.write("AL8", "Progress", label_fmt)

    for i, _ in enumerate(habit_names):
        analysis_row = 8 + i  # 0‑based -> Excel rows 9–...
        habit_excel_row = habit_start_row + 1 + i  # linked to habit name

        # Habit name (linked)
        habits_ws.write_formula(analysis_row, 34, f"=B{habit_excel_row}")

        # Goal = number of days in the month
        goal_formula = f"=COUNTA(C{daynum_row+1}:{last_col_letter}{daynum_row+1})"
        habits_ws.write_formula(analysis_row, 35, goal_formula)

        # Actual completions
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
    # WEEK SHEET  (monthly planning with horizontal row colors)
    # ------------------------------------------------------------------
    week_ws = workbook.add_worksheet("Week")

    week_ws.set_column("A:A", 2)
    for col in ["B", "D", "F", "H", "J", "L", "N"]:
        week_ws.set_column(f"{col}:{col}", 3)   # checkbox column
    for col in ["C", "E", "G", "I", "K", "M", "O"]:
        week_ws.set_column(f"{col}:{col}", 22)  # task text column
    week_ws.set_column("Q:R", 16)

    # Title
    week_ws.merge_range("B2:O4", "Weekly Planning", title_fmt)

    # We layout whole month as weekly blocks stacked vertically
    start_cols = [1, 3, 5, 7, 9, 11, 13]  # B, D, F, H, J, L, N

    tasks_per_day = 10
    block_height = tasks_per_day + 7  # date + day + header + tasks + 3 stats + blank
    first_block_top = 8  # row index (0‑based) for first week's date row (Excel row 9)

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
            wd = date_obj.weekday()
            day_name = date_obj.strftime("%a").upper()

            # Date row (day of month)
            rng_date = (
                f"{xl_rowcol_to_cell(row_day_date, tick_col)}:"
                f"{xl_rowcol_to_cell(row_day_date, text_col)}"
            )
            if wd == 6:
                week_ws.merge_range(rng_date, date_obj.day, week_date_fmt_sun)
            else:
                week_ws.merge_range(rng_date, date_obj.day, week_date_fmt_normal)

            # Day name row (Sunday red)
            rng_title = (
                f"{xl_rowcol_to_cell(row_day_title, tick_col)}:"
                f"{xl_rowcol_to_cell(row_day_title, text_col)}"
            )
            if wd == 6:
                week_ws.merge_range(rng_title, day_name, week_day_header_sun_fmt)
            else:
                week_ws.merge_range(rng_title, day_name, week_day_header_fmt)

            # Tasks header row
            rng_tasks = (
                f"{xl_rowcol_to_cell(row_task_header, tick_col)}:"
                f"{xl_rowcol_to_cell(row_task_header, text_col)}"
            )
            if wd == 6:
                week_ws.merge_range(rng_tasks, "Tasks", week_tasks_header_sun_fmt)
            else:
                week_ws.merge_range(rng_tasks, "Tasks", week_tasks_header_fmt)

            # Tasks area (rows colored horizontally)
            top_tick = row_task_start
            bottom_tick = row_task_start + tasks_per_day - 1
            for r in range(top_tick, bottom_tick + 1):
                row_idx = r - row_task_start  # 0..tasks_per_day-1
                color_idx = row_idx % len(week_row_colors)

                week_ws.data_validation(
                    r,
                    tick_col,
                    r,
                    tick_col,
                    {"validate": "list", "source": ["☐", "☑"]},
                )
                week_ws.write(r, tick_col, "☐", week_tick_row_fmts[color_idx])
                week_ws.write(r, text_col, "", week_text_row_fmts[color_idx])

            tick_top_addr = xl_rowcol_to_cell(top_tick, tick_col)
            tick_bot_addr = xl_rowcol_to_cell(bottom_tick, tick_col)
            task_top_addr = xl_rowcol_to_cell(top_tick, text_col)
            task_bot_addr = xl_rowcol_to_cell(bottom_tick, text_col)

            # Stats rows (neutral formatting; not row‑colored)
            week_ws.write(row_completed, tick_col, "Completed", label_fmt)
            week_ws.write(
                row_completed,
                text_col,
                f'=COUNTIF({tick_top_addr}:{tick_bot_addr},"☑")',
                border_center_fmt,
            )

            week_ws.write(row_notcompleted, tick_col, "Not Completed", label_fmt)
            week_ws.write(
                row_notcompleted,
                text_col,
                f'=COUNTIF({tick_top_addr}:{tick_bot_addr},"☐")',
                border_center_fmt,
            )

            week_ws.write(row_total, tick_col, "Total tasks", label_fmt)
            week_ws.write(
                row_total,
                text_col,
                f"=COUNTA({task_top_addr}:{task_bot_addr})",
                border_center_fmt,
            )

            day_counter += 1

    workbook.close()


if __name__ == "__main__":
    create_focus_tracker()