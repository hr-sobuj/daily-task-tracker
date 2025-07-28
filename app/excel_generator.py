import xlsxwriter
import calendar
from datetime import date, timedelta

def generate_task_tracker(year: int, filename: str):
    # ===== User-defined Goals =====
    WEEKLY_GOAL = 20
    MONTHLY_GOAL = 80
    YEARLY_GOAL = 1000
    DAILY_GOAL = int(WEEKLY_GOAL / 7)  # tasks per day

    workbook = xlsxwriter.Workbook(filename)

    # ===== Formats =====
    header_format = workbook.add_format({
        'bold': True, 'font_color': 'white', 'bg_color': '#0B5394',
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    date_format     = workbook.add_format({'num_format': 'dd-mm-yyyy', 'border': 1, 'align': 'center'})
    centered_format = workbook.add_format({'border': 1, 'align': 'center'})
    done_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    pend_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
    skip_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    high_fmt   = workbook.add_format({'bg_color': '#FFCDD2'})
    medium_fmt = workbook.add_format({'bg_color': '#FFF9C4'})
    low_fmt    = workbook.add_format({'bg_color': '#C8E6C9'})
    time_format = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'center'})
    percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})
    cell_format = workbook.add_format({'border': 1, 'align': 'center'})

    # ===== Goals Sheet =====
    goals = workbook.add_worksheet("Goals")
    goals.write_row(0, 0, ["Goal (Tasks)", "Value"], header_format)
    goals.write(1, 0, "Daily");   goals.write_number(1, 1, DAILY_GOAL)
    goals.write(2, 0, "Weekly");  goals.write_number(2, 1, WEEKLY_GOAL)
    goals.write(3, 0, "Monthly"); goals.write_number(3, 1, MONTHLY_GOAL)
    goals.write(4, 0, "Yearly");  goals.write_number(4, 1, YEARLY_GOAL)
    goals.write(5, 0, "Unit");    goals.write(5, 1, "Tasks")
    goals.set_column(0, 0, 20)
    goals.set_column(1, 1, 12)

    # ===== Monthly Sheets =====
    task_headers = [
        "Date", "Day", "Daily Goal", "Task Description", "Priority", "Status",
        "Hours Spent", "What Went Well", "What I Missed", "Notes"
    ]
    col_widths = [15, 15, 12, 30, 12, 15, 12, 25, 25, 30]
    priority_list = ['High', 'Medium', 'Low']
    status_list   = ['Pending', 'Done', 'Skipped']

    for month in range(1, 13):
        name = calendar.month_name[month]
        sheet = workbook.add_worksheet(name)
        days = calendar.monthrange(year, month)[1]
        sheet.write_row(0, 0, task_headers, header_format)
        for col, width in enumerate(col_widths):
            sheet.set_column(col, col, width)
        for day in range(1, days + 1):
            current = date(year, month, day)
            sheet.write_datetime(day, 0, current, date_format)
            sheet.write(day, 1, current.strftime('%A'), centered_format)
            # এখানে এখন formula use করলাম যা Goals শীট থেকে daily goal নিবে
            sheet.write_formula(day, 2, '=Goals!B2', centered_format)
            for col in range(3, len(task_headers)):
                if col == 6:
                    sheet.write(day, col, None, time_format)
                else:
                    sheet.write(day, col, None, centered_format)

        # Dropdown validation
        sheet.data_validation(f'E2:E{days+1}', {'validate': 'list', 'source': priority_list})
        sheet.data_validation(f'F2:F{days+1}', {'validate': 'list', 'source': status_list})

        # Conditional formats
        sheet.conditional_format(f'E2:E{days+1}', {'type': 'text', 'criteria': 'containing', 'value': 'High', 'format': high_fmt})
        sheet.conditional_format(f'E2:E{days+1}', {'type': 'text', 'criteria': 'containing', 'value': 'Medium', 'format': medium_fmt})
        sheet.conditional_format(f'E2:E{days+1}', {'type': 'text', 'criteria': 'containing', 'value': 'Low', 'format': low_fmt})
        sheet.conditional_format(f'F2:F{days+1}', {'type': 'text', 'criteria': 'containing', 'value': 'Done', 'format': done_fmt})
        sheet.conditional_format(f'F2:F{days+1}', {'type': 'text', 'criteria': 'containing', 'value': 'Pending', 'format': pend_fmt})
        sheet.conditional_format(f'F2:F{days+1}', {'type': 'text', 'criteria': 'containing', 'value': 'Skipped', 'format': skip_fmt})

    # ===== Weekly Report =====
    weekly = workbook.add_worksheet("Weekly Report")
    weekly_headers = ["Week No", "Total Tasks", "Done", "Pending", "Skipped", "Total Hours", "Goal", "% Complete"]
    weekly.write_row(0, 0, weekly_headers, header_format)
    weekly.set_column(0, 7, 15)
    weekly.write_comment(0, 6, f"Weekly goal: linked from Goals sheet")

    for w in range(1, 53):
        weekly.write_number(w, 0, w)
        start = date(year, 1, 1) + timedelta(days=(w - 1) * 7)
        end   = start + timedelta(days=6)
        start_str = start.strftime("%Y-%m-%d")
        end_str   = end.strftime("%Y-%m-%d")

        total_parts, done_parts, pend_parts, skip_parts, hr_parts = [], [], [], [], []

        for m in range(1, 13):
            mn = calendar.month_name[m]
            last = calendar.monthrange(year, m)[1]
            date_rng   = f"'{mn}'!$A$2:$A${last+1}"
            desc_rng   = f"'{mn}'!$D$2:$D${last+1}"
            status_rng = f"'{mn}'!$F$2:$F${last+1}"
            hours_rng  = f"'{mn}'!$G$2:$G${last+1}"

            total_parts.append(
                f"SUMPRODUCT((--({date_rng}>=DATEVALUE(\"{start_str}\")))*"
                f"(--({date_rng}<=DATEVALUE(\"{end_str}\")))*(LEN({desc_rng})>0))"
            )
            done_parts.append(
                f"SUMPRODUCT((--({date_rng}>=DATEVALUE(\"{start_str}\")))*"
                f"(--({date_rng}<=DATEVALUE(\"{end_str}\")))*(LEN({desc_rng})>0)*"
                f"({status_rng}=\"Done\"))"
            )
            pend_parts.append(
                f"SUMPRODUCT((--({date_rng}>=DATEVALUE(\"{start_str}\")))*"
                f"(--({date_rng}<=DATEVALUE(\"{end_str}\")))*(LEN({desc_rng})>0)*"
                f"({status_rng}=\"Pending\"))"
            )
            skip_parts.append(
                f"SUMPRODUCT((--({date_rng}>=DATEVALUE(\"{start_str}\")))*"
                f"(--({date_rng}<=DATEVALUE(\"{end_str}\")))*(LEN({desc_rng})>0)*"
                f"({status_rng}=\"Skipped\"))"
            )
            hr_parts.append(
                f"SUMPRODUCT((--({date_rng}>=DATEVALUE(\"{start_str}\")))*"
                f"(--({date_rng}<=DATEVALUE(\"{end_str}\")))*(--({hours_rng}))*1)"
            )

        weekly.write_formula(w, 1, "=" + "+".join(total_parts), cell_format)
        weekly.write_formula(w, 2, "=" + "+".join(done_parts), cell_format)
        weekly.write_formula(w, 3, "=" + "+".join(pend_parts), cell_format)
        weekly.write_formula(w, 4, "=" + "+".join(skip_parts), cell_format)
        weekly.write_formula(w, 5, "=" + "+".join(hr_parts), time_format)
        # Dynamic link to Goals sheet for Weekly Goal
        weekly.write_formula(w, 6, '=Goals!B3', cell_format)
        weekly.write_formula(w, 7, f'=C{w+1}/G{w+1}', percent_format)

    weekly.conditional_format('H2:H53', {'type': 'cell', 'criteria': '>=', 'value': 1, 'format': done_fmt})
    weekly.conditional_format('H2:H53', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': pend_fmt})

    # ===== Monthly Report =====
    monthly_headers = ["Month Name", "Total Tasks", "Done", "Pending", "Skipped", "Total Hours", "Goal", "% Complete"]
    monthly = workbook.add_worksheet("Monthly Report")
    monthly.write_row(0, 0, monthly_headers, header_format)
    monthly.set_column(0, 7, 15)
    monthly.write_comment(0, 6, f"Monthly goal: linked from Goals sheet")

    for m in range(1, 13):
        name = calendar.month_name[m]
        last = calendar.monthrange(year, m)[1]
        monthly.write(m, 0, name, cell_format)
        monthly.write_formula(m, 1, f"=COUNTIF('{name}'!$D$2:$D${last+1},\"<>\")", cell_format)
        monthly.write_formula(m, 2, f"=COUNTIFS('{name}'!$D$2:$D${last+1},\"<>\",'{name}'!$F$2:$F${last+1},\"Done\")", cell_format)
        monthly.write_formula(m, 3, f"=COUNTIFS('{name}'!$D$2:$D${last+1},\"<>\",'{name}'!$F$2:$F${last+1},\"Pending\")", cell_format)
        monthly.write_formula(m, 4, f"=COUNTIFS('{name}'!$D$2:$D${last+1},\"<>\",'{name}'!$F$2:$F${last+1},\"Skipped\")", cell_format)
        monthly.write_formula(m, 5, f"=SUM('{name}'!$G$2:$G${last+1})", time_format)
        # Dynamic link to Goals sheet for Monthly Goal
        monthly.write_formula(m, 6, '=Goals!B4', cell_format)
        monthly.write_formula(m, 7, f'=C{m+1}/G{m+1}', percent_format)

    monthly.conditional_format('H2:H13', {'type': 'cell', 'criteria': '>=', 'value': 1, 'format': done_fmt})
    monthly.conditional_format('H2:H13', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': pend_fmt})

    # ===== Yearly Report =====
    yearly_headers = ["Year", "Total Tasks", "Done", "Pending", "Skipped", "Total Hours", "Goal", "% Complete"]
    yearly = workbook.add_worksheet("Yearly Report")
    yearly.write_row(0, 0, yearly_headers, header_format)
    yearly.set_column(0, 7, 15)

    # Calculate full year total
    total_parts, done_parts, pend_parts, skip_parts, hr_parts = [], [], [], [], []

    for m in range(1, 13):
        mn = calendar.month_name[m]
        last = calendar.monthrange(year, m)[1]
        date_rng   = f"'{mn}'!$A$2:$A${last+1}"
        desc_rng   = f"'{mn}'!$D$2:$D${last+1}"
        status_rng = f"'{mn}'!$F$2:$F${last+1}"
        hours_rng  = f"'{mn}'!$G$2:$G${last+1}"

        total_parts.append(f"COUNTIF('{mn}'!$D$2:$D${last+1},\"<>\")")
        done_parts.append(f"COUNTIFS('{mn}'!$D$2:$D${last+1},\"<>\",'{mn}'!$F$2:$F${last+1},\"Done\")")
        pend_parts.append(f"COUNTIFS('{mn}'!$D$2:$D${last+1},\"<>\",'{mn}'!$F$2:$F${last+1},\"Pending\")")
        skip_parts.append(f"COUNTIFS('{mn}'!$D$2:$D${last+1},\"<>\",'{mn}'!$F$2:$F${last+1},\"Skipped\")")
        hr_parts.append(f"SUM('{mn}'!$G$2:$G${last+1})")

    yearly.write_number(1, 0, year)
    yearly.write_formula(1, 1, "=" + "+".join(total_parts), cell_format)
    yearly.write_formula(1, 2, "=" + "+".join(done_parts), cell_format)
    yearly.write_formula(1, 3, "=" + "+".join(pend_parts), cell_format)
    yearly.write_formula(1, 4, "=" + "+".join(skip_parts), cell_format)
    yearly.write_formula(1, 5, "=" + "+".join(hr_parts), time_format)
    # Dynamic link to Goals sheet for Yearly Goal
    yearly.write_formula(1, 6, '=Goals!B5', cell_format)
    yearly.write_formula(1, 7, '=C2/G2', percent_format)

    yearly.conditional_format('H2:H2', {'type': 'cell', 'criteria': '>=', 'value': 1, 'format': done_fmt})
    yearly.conditional_format('H2:H2', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': pend_fmt})

    workbook.close()
