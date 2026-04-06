# Multi-Line Production Scheduler

A Web-Based Production Scheduler built with Streamlit and Python.

## Features
- **Waterfall Scheduling**: Automatically schedules parts sequentially based on line capacity.
- **Setup Management**: Handles Major (240 min) and Minor (60 min) setup times.
- **Calendar Awareness**: Schedules for January 2026, automatically skipping Sundays.
- **Excel Integration**: Upload your plan and download the fully scheduled result.

## Setup & Run

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the App**:
   ```bash
   streamlit run app.py
   ```

## Input File Format
The Excel file must have the following columns (starting at Row 1):
- `Sr. No.`
- `Line Name` (e.g., Arjun-1, AutoLine)
- `Part Number`
- `Part Description`
- `Total Plan Qty`
- `Major Setup` (1 for Yes)
- `Minor Setup` (1 for Yes)
