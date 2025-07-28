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
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 16,
        'align': 'left',
        'valign': 'vcenter',
        'font_color': '#0B5394'
    })
    header_format = workbook.add_format({
        'bold': True, 'font_color': 'white', 'bg_color': '#0B5394',
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    date_format = workbook.add_format({'num_format': 'dd-mm-yyyy', 'border': 1, 'align': 'center'})
    centered_format = workbook.add_format({'border': 1, 'align': 'center'})
    done_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    pend_fmt = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
    skip_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    high_fmt = workbook.add_format({'bg_color': '#FFCDD2'})
    medium_fmt = workbook.add_format({'bg_color': '#FFF9C4'})
    low_fmt = workbook.add_format({'bg_color': '#C8E6C9'})
    time_format = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'center'})
    percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1, 'align': 'center'})
    cell_format = workbook.add_format({'border': 1, 'align': 'center'})

    # ===== Goals Sheet =====
    goals = workbook.add_worksheet("Goals")
    goals.merge_range('A1:B1', f'Goals Overview {year}', title_format)
    goals.write_row(2, 0, ["Goal (Tasks)", "Value"], header_format)
    goals.write(3, 0, "Daily"); goals.write_number(3, 1, DAILY_GOAL)
    goals.write(4, 0, "Weekly"); goals.write_number(4, 1, WEEKLY_GOAL)
    goals.write(5, 0, "Monthly"); goals.write_number(5, 1, MONTHLY_GOAL)
    goals.write(6, 0, "Yearly"); goals.write_number(6, 1, YEARLY_GOAL)
    goals.write(7, 0, "Unit"); goals.write(7, 1, "Tasks")
    goals.set_column(0, 0, 20)
    goals.set_column(1, 1, 12)

    # ===== Monthly Sheets =====
    task_headers = [
        "Date", "Day", "Daily Goal", "Task Description", "Priority", "Status",
        "Hours Spent", "What Went Well", "What I Missed", "Notes"
    ]
    col_widths = [15, 15, 12, 30, 12, 15, 12, 25, 25, 30]
    priority_list = ['High', 'Medium', 'Low']
    status_list = ['Pending', 'Done', 'Skipped']

    for month in range(1, 13):
        name = calendar.month_name[month]
        sheet = workbook.add_worksheet(name)
        days = calendar.monthrange(year, month)[1]
        sheet.merge_range(0, 0, 0, len(task_headers) - 1, f'{name} {year} Daily Tasks', title_format)
        sheet.write_row(1, 0, task_headers, header_format)
        for col, width in enumerate(col_widths):
            sheet.set_column(col, col, width)
        for day in range(1, days + 1):
            current = date(year, month, day)
            sheet.write_datetime(day + 1, 0, current, date_format)
            sheet.write(day + 1, 1, current.strftime('%A'), centered_format)
            sheet.write_formula(day + 1, 2, '=Goals!B4', centered_format)
            for col in range(3, len(task_headers)):
                fmt = time_format if col == 6 else centered_format
                sheet.write(day + 1, col, None, fmt)
        sheet.data_validation(f'E3:E{days+2}', {'validate': 'list', 'source': priority_list})
        sheet.data_validation(f'F3:F{days+2}', {'validate': 'list', 'source': status_list})
        sheet.conditional_format(f'E3:E{days+2}', {'type': 'text', 'criteria': 'containing', 'value': 'High', 'format': high_fmt})
        sheet.conditional_format(f'E3:E{days+2}', {'type': 'text', 'criteria': 'containing', 'value': 'Medium', 'format': medium_fmt})
        sheet.conditional_format(f'E3:E{days+2}', {'type': 'text', 'criteria': 'containing', 'value': 'Low', 'format': low_fmt})
        sheet.conditional_format(f'F3:F{days+2}', {'type': 'text', 'criteria': 'containing', 'value': 'Done', 'format': done_fmt})
        sheet.conditional_format(f'F3:F{days+2}', {'type': 'text', 'criteria': 'containing', 'value': 'Pending', 'format': pend_fmt})
        sheet.conditional_format(f'F3:F{days+2}', {'type': 'text', 'criteria': 'containing', 'value': 'Skipped', 'format': skip_fmt})

    # ===== Weekly Report =====
    weekly = workbook.add_worksheet("Weekly Report")
    weekly_headers = ["Week No", "Total Tasks", "Done", "Pending", "Skipped", "Total Hours", "Goal", "% Complete"]
    weekly.merge_range(0, 0, 0, len(weekly_headers) - 1, f'Weekly Task Report {year}', title_format)
    weekly.write_row(1, 0, weekly_headers, header_format)
    weekly.set_column(0, len(weekly_headers) - 1, 15)

    for w in range(1, 53):
        row = w + 1
        weekly.write_number(row, 0, w, cell_format)
        start = date(year, 1, 1) + timedelta(days=(w - 1) * 7)
        end = start + timedelta(days=6)
        start_str = start.strftime("%Y-%m-%d")
        end_str = end.strftime("%Y-%m-%d")

        total_parts = []
        done_parts = []
        pend_parts = []
        skip_parts = []
        hrs_parts = []
        for m in range(1, 13):
            mn = calendar.month_name[m]
            last = calendar.monthrange(year, m)[1]
            date_rng = f"'{mn}'!$A$2:$A${last+1}"
            desc_rng = f"'{mn}'!$D$2:$D${last+1}"
            status_rng = f"'{mn}'!$F$2:$F${last+1}"
            hours_rng = f"'{mn}'!$G$2:$G${last+1}"

            total_parts.append(
                f'COUNTIFS({date_rng}, ">="&DATEVALUE("{start_str}"), {date_rng}, "<="&DATEVALUE("{end_str}"), {desc_rng}, "<>")'
            )
            done_parts.append(
                f'COUNTIFS({date_rng}, ">="&DATEVALUE("{start_str}"), {date_rng}, "<="&DATEVALUE("{end_str}"), {desc_rng}, "<>", {status_rng}, "Done")'
            )
            pend_parts.append(
                f'COUNTIFS({date_rng}, ">="&DATEVALUE("{start_str}"), {date_rng}, "<="&DATEVALUE("{end_str}"), {desc_rng}, "<>", {status_rng}, "Pending")'
            )
            skip_parts.append(
                f'COUNTIFS({date_rng}, ">="&DATEVALUE("{start_str}"), {date_rng}, "<="&DATEVALUE("{end_str}"), {desc_rng}, "<>", {status_rng}, "Skipped")'
            )
            hrs_parts.append(
                f'SUMIFS({hours_rng}, {date_rng}, ">="&DATEVALUE("{start_str}"), {date_rng}, "<="&DATEVALUE("{end_str}"))'
            )

        weekly.write_formula(row, 1, "=" + "+".join(total_parts), cell_format)
        weekly.write_formula(row, 2, "=" + "+".join(done_parts), cell_format)
        weekly.write_formula(row, 3, "=" + "+".join(pend_parts), cell_format)
        weekly.write_formula(row, 4, "=" + "+".join(skip_parts), cell_format)
        weekly.write_formula(row, 5, "=" + "+".join(hrs_parts), time_format)
        weekly.write_formula(row, 6, '=Goals!B5', cell_format)
        weekly.write_formula(row, 7, f'=IF(G{row+1}=0, 0, C{row+1}/G{row+1})', percent_format)
    weekly.conditional_format('H3:H54', {'type': 'cell', 'criteria': '>=', 'value': 1, 'format': done_fmt})
    weekly.conditional_format('H3:H54', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': pend_fmt})

    # ===== Monthly Report =====
    monthly_headers = ["Month Name", "Total Tasks", "Done", "Pending", "Skipped", "Total Hours", "Goal", "% Complete"]
    monthly = workbook.add_worksheet("Monthly Report")
    monthly.merge_range(0, 0, 0, len(monthly_headers) - 1, f'Monthly Task Report {year}', title_format)
    monthly.write_row(1, 0, monthly_headers, header_format)
    monthly.set_column(0, len(monthly_headers) - 1, 15)

    for m in range(1, 13):
        row = m + 1
        name = calendar.month_name[m]
        last = calendar.monthrange(year, m)[1]
        desc_rng = f"'{name}'!$D$3:$D${last+2}"
        status_rng = f"'{name}'!$F$3:$F${last+2}"
        hours_rng = f"'{name}'!$G$3:$G${last+2}"

        monthly.write(row, 0, name, cell_format)
        monthly.write_formula(row, 1, f'=COUNTIF({desc_rng}, "<>")', cell_format)
        monthly.write_formula(row, 2, f'=COUNTIFS({desc_rng}, "<>", {status_rng}, "Done")', cell_format)
        monthly.write_formula(row, 3, f'=COUNTIFS({desc_rng}, "<>", {status_rng}, "Pending")', cell_format)
        monthly.write_formula(row, 4, f'=COUNTIFS({desc_rng}, "<>", {status_rng}, "Skipped")', cell_format)
        monthly.write_formula(row, 5, f'=SUM({hours_rng})', time_format)
        monthly.write_formula(row, 6, '=Goals!B6', cell_format)
        monthly.write_formula(row, 7, f'=IF(G{row+1}=0, 0, C{row+1}/G{row+1})', percent_format)
    monthly.conditional_format('H3:H14', {'type': 'cell', 'criteria': '>=', 'value': 1, 'format': done_fmt})
    monthly.conditional_format('H3:H14', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': pend_fmt})

    # ===== Yearly Report =====
    yearly_headers = ["Year", "Total Tasks", "Done", "Pending", "Skipped", "Total Hours", "Goal", "% Complete"]
    yearly = workbook.add_worksheet("Yearly Report")
    yearly.merge_range(0, 0, 0, len(yearly_headers) - 1, f'Yearly Task Report {year}', title_format)
    yearly.write_row(1, 0, yearly_headers, header_format)
    yearly.set_column(0, len(yearly_headers) - 1, 15)

    total_parts, done_parts, pend_parts, skip_parts, hr_parts = [], [], [], [], []
    for m in range(1, 13):
        mn = calendar.month_name[m]
        last = calendar.monthrange(year, m)[1]
        desc_rng = f"'{mn}'!$D$3:$D${last+2}"
        status_rng = f"'{mn}'!$F$3:$F${last+2}"
        hours_rng = f"'{mn}'!$G$3:$G${last+2}"

        total_parts.append(f'COUNTIF({desc_rng}, "<>")')
        done_parts.append(f'COUNTIFS({desc_rng}, "<>", {status_rng}, "Done")')
        pend_parts.append(f'COUNTIFS({desc_rng}, "<>", {status_rng}, "Pending")')
        skip_parts.append(f'COUNTIFS({desc_rng}, "<>", {status_rng}, "Skipped")')
        hr_parts.append(f'SUM({hours_rng})')
    yearly.write_number(2, 0, year)
    yearly.write_formula(2, 1, "=" + "+".join(total_parts), cell_format)
    yearly.write_formula(2, 2, "=" + "+".join(done_parts), cell_format)
    yearly.write_formula(2, 3, "=" + "+".join(pend_parts), cell_format)
    yearly.write_formula(2, 4, "=" + "+".join(skip_parts), cell_format)
    yearly.write_formula(2, 5, "=" + "+".join(hr_parts), time_format)
    yearly.write_formula(2, 6, '=Goals!B7', cell_format)
    yearly.write_formula(2, 7, '=IF(G3=0, 0, C3/G3)', percent_format)
    yearly.conditional_format('H3:H3', {'type': 'cell', 'criteria': '>=', 'value': 1, 'format': done_fmt})
    yearly.conditional_format('H3:H3', {'type': 'cell', 'criteria': '<', 'value': 1, 'format': pend_fmt})

    workbook.close()

# Example usage:
# generate_task_tracker(2025, 'Task_Tracker_2025.xlsx')
