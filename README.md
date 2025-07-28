# Daily Task Tracker

A Python-based application that generates Excel workbooks for tracking daily tasks, with weekly, monthly, and yearly progress reports.

## Features

- Generates task tracking Excel workbooks with multiple sheets
- Pre-defined goals tracking (Daily, Weekly, Monthly, Yearly)
- Task priority levels (High, Medium, Low)
- Task status tracking (Done, Pending, Skipped)
- Detailed reports:
  - Weekly Report
  - Monthly Report
  - Yearly Report
- Color-coded status and priority indicators
- Time tracking for tasks
- Progress percentage calculations

## Requirements

- Python 3.x
- XlsxWriter library

## Installation

1. Clone the repository:
```bash
git clone https://github.com/hr-sobuj/daily-task-tracker.git
cd daily-task-tracker
```

2. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the main script:
```bash
python main.py
```

2. The program will generate Excel files for task tracking (e.g., task_tracker_2026.xlsx)

## Excel Workbook Structure

The generated Excel workbook contains the following sheets:

1. **Goals Sheet**: Contains predefined goals for daily, weekly, monthly, and yearly tasks
2. **Monthly Sheets**: 12 sheets (January-December) for daily task entries with:
   - Date and day tracking
   - Task descriptions
   - Priority levels
   - Status tracking
   - Time tracking
   - Notes and feedback sections
3. **Weekly Report**: Summarizes task completion statistics by week
4. **Monthly Report**: Provides monthly task completion overview
5. **Yearly Report**: Shows yearly progress and statistics

## Contributing

Feel free to submit issues and enhancement requests.
