import streamlit as st
import pandas as pd
import datetime
import numpy as np
from io import BytesIO

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# --- Configuration & Constants ---
DAILY_MINUTES = 1320
CAPACITIES = {
    "Arjun-1": 10117,
    "Arjun-2": 12500,
    "Arjun-3": 12500,
    "Arjun-4": 11800,
    "Arjun-5": 10000,
    "Arjun-6": 28089,
    "Arjun-7": 7920,
    "Arjun-8": 9900,
    "Arjun-9": 12500,
    "Arjun-10": 6500,
    "Arjun-11": 9240,
    "Arjun-12": 8088,
    "AutoLine": 24568,
}

# --- Core Logic ---
def get_monthly_working_days(start_date, end_date):
    """Generates dates, skipping Sundays."""
    dates = []
    current = start_date
    while current <= end_date:
        if current.weekday() != 6:  # 6 is Sunday
            dates.append(current)
        current += datetime.timedelta(days=1)
    return dates

def get_weekly_working_days(start_date, num_days=6):
    """Generates exactly `num_days` working days starting from start_date, skipping Sundays."""
    dates = []
    current = start_date
    while len(dates) < num_days:
        if current.weekday() != 6:
            dates.append(current)
        current += datetime.timedelta(days=1)
    return dates

def process_schedule(df, working_days):
    required_cols = [
        "Sr. No.", "Line Name", "Part Number", "Part Description", 
        "Total Plan Qty", "Major Setup", "Minor Setup"
    ]
    
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns in uploaded file: {', '.join(missing)}")
        return None

    # Create columns for all dates if they don't exist
    for d in working_days:
        if d not in df.columns:
            df[d] = 0

    lines = df["Line Name"].unique()
    line_states = {} 
    
    for line in lines:
        if line in CAPACITIES:
            line_states[line] = {"day_idx": 0, "remaining_cap": CAPACITIES[line]}
        else:
            line_states[line] = {"day_idx": 0, "remaining_cap": 0}

    for index, row in df.iterrows():
        line_name = row["Line Name"]
        if line_name not in CAPACITIES:
            continue
            
        daily_cap = CAPACITIES[line_name]
        units_per_min = daily_cap / DAILY_MINUTES
        
        state = line_states[line_name]
        day_idx = state["day_idx"]
        current_cap = state["remaining_cap"]
        
        major_setup = row.get("Major Setup", 0)
        minor_setup = row.get("Minor Setup", 0)
        
        deduction_units = 0
        if major_setup == 1:
            deduction_units += (units_per_min * 240)
        if minor_setup == 1:
            deduction_units += (units_per_min * 60)
            
        current_cap -= deduction_units
        
        while current_cap < 0:
            day_idx += 1
            if day_idx >= len(working_days):
                break 
            debt = abs(current_cap)
            current_cap = daily_cap - debt
        
        state["day_idx"] = day_idx
        state["remaining_cap"] = current_cap
        
        if day_idx >= len(working_days):
            continue 

        qty_needed = row["Total Plan Qty"]
        
        while qty_needed > 0:
            if day_idx >= len(working_days):
                break
                
            if current_cap < 1:
                day_idx += 1
                if day_idx >= len(working_days):
                    break
                current_cap = daily_cap

            can_make = int(current_cap)
            make_today = min(qty_needed, can_make)
            
            date_col = working_days[day_idx]
            current_val = df.at[index, date_col]
            df.at[index, date_col] = current_val + make_today
            
            qty_needed -= make_today
            current_cap -= make_today

        line_states[line_name] = {
            "day_idx": day_idx,
            "remaining_cap": current_cap
        }

    return df

def get_template_buffer():
    # Create Template
    template_cols = ["Sr. No.", "Line Name", "Part Number", "Part Description", 
                     "Total Plan Qty", "Major Setup", "Minor Setup"]
    sample_data = [[1, "Arjun-1", "SAMPLE-001", "Example Part", 5000, 1, 0]]
    template_df = pd.DataFrame(sample_data, columns=template_cols)
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Input_Template')
    return buffer

def generate_monthly_excel(result_df, final_cols, title):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        sheet_name = 'Schedule'
        worksheet = workbook.add_worksheet(sheet_name)
        
        main_title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        line_title_fmt = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFFFF', 'border': 1})
        header_fmt = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        data_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        data_left_fmt = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        date_head_fmt = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78', 'border': 1, 'num_format': 'dd-mmm', 'align': 'center', 'valign': 'vcenter'})

        n_cols = len(final_cols)
        worksheet.set_column(0, 0, 5)   # Sr No
        worksheet.set_column(1, 1, 10)  # Line
        worksheet.set_column(2, 2, 15)  # Part No
        worksheet.set_column(3, 3, 30)  # Desc
        worksheet.set_column(4, n_cols-1, 8) 

        current_row = 0
        worksheet.merge_range(current_row, 0, current_row, n_cols-1, title, main_title_fmt)
        current_row += 1
        
        unique_lines = result_df["Line Name"].unique()
        for line in unique_lines:
            line_df = result_df[result_df["Line Name"] == line][final_cols]
            worksheet.merge_range(current_row, 0, current_row, n_cols-1, f"{line} Production Plan", line_title_fmt)
            current_row += 1
            
            for col_idx, col_name in enumerate(final_cols):
                if isinstance(col_name, datetime.date):
                    worksheet.write(current_row, col_idx, col_name, date_head_fmt)
                else:
                    worksheet.write(current_row, col_idx, str(col_name), header_fmt)
            current_row += 1
            
            sr_no_counter = 1
            for _, row_data in line_df.iterrows():
                for col_idx, col_name in enumerate(final_cols):
                    val = sr_no_counter if col_idx == 0 else row_data[col_name]
                    cell_fmt = data_left_fmt if col_idx in [2, 3] else data_fmt
                    if pd.isna(val):
                        worksheet.write(current_row, col_idx, "", cell_fmt)
                    else:
                        worksheet.write(current_row, col_idx, val, cell_fmt)
                current_row += 1
                sr_no_counter += 1
            
            worksheet.write(current_row, 0, "", total_fmt)
            worksheet.write(current_row, 1, "", total_fmt)
            worksheet.write(current_row, 2, "Total", total_fmt)
            worksheet.write(current_row, 3, "", total_fmt)
            
            for col_idx in range(4, n_cols):
                total_val = line_df[final_cols[col_idx]].sum()
                worksheet.write(current_row, col_idx, total_val, total_fmt)
            
            current_row += 3
    return output.getvalue()


def generate_weekly_excel_report(result_df, final_cols, start_date, selected_week):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        sheet_name = 'Weekly_Schedule'
        worksheet = workbook.add_worksheet(sheet_name)
        
        # Apply Print Layout and Margins
        worksheet.set_landscape()
        worksheet.set_paper(9) # A4
        worksheet.fit_to_pages(1, 0) # Fit width to 1 page
        worksheet.set_margins(left=0.25, right=0.25, top=0.5, bottom=0.5)
        worksheet.set_default_row(22) # Globally set row height to 22

        # --- Formats ---
        company_title_fmt = workbook.add_format({
            'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        format_no_fmt = workbook.add_format({
            'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        month_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        week_num_fmt = workbook.add_format({
            'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        line_title_fmt = workbook.add_format({
            'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        header_fmt = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78', 
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        downtime_title_fmt = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78', 
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        data_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        data_left_fmt = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
        total_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        date_head_fmt = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78', 
            'border': 1, 'num_format': 'dd-mmm', 'align': 'center', 'valign': 'vcenter'
        })
        actual_data_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        
        display_cols = ["Sr No", "Part Number", "Part Description", "Weekly Plan Qty"]
        date_cols = [c for c in final_cols if isinstance(c, datetime.date)]
        weekly_output_cols = display_cols + date_cols
        n_weekly_cols = len(weekly_output_cols)

        worksheet.set_column(0, 0, 5)   # Sr No
        worksheet.set_column(1, 1, 18)  # Part Number
        worksheet.set_column(2, 2, 40)  # Part Description
        worksheet.set_column(3, 3, 14)  # Weekly Plan Qty
        worksheet.set_column(4, n_weekly_cols-1, 10) # Date cols

        current_row = 0
        h_pagebreaks = []
        
        month_title = f"Weekly Production Plan {start_date.strftime('%B %Y')}"

        unique_lines = result_df["Line Name"].unique()
        
        for idx, line in enumerate(unique_lines):
            line_df = result_df[result_df["Line Name"] == line]
            
            # --- ROW 0: Company Info & Format No ---
            worksheet.merge_range(current_row, 0, current_row, n_weekly_cols-3, "KSPG Automotive India Pvt Ltd , BU Bearing", company_title_fmt)
            worksheet.merge_range(current_row, n_weekly_cols-2, current_row, n_weekly_cols-1, "Format No:- PRH.F.46.00", format_no_fmt)
            current_row += 1
            
            # --- ROW 1: Month Info & Week No ---
            worksheet.merge_range(current_row, 0, current_row, n_weekly_cols-3, month_title, month_title_fmt)
            worksheet.merge_range(current_row, n_weekly_cols-2, current_row, n_weekly_cols-1, selected_week, week_num_fmt)
            current_row += 1
            
            # --- PLAN SECTION ---
            worksheet.merge_range(current_row, 0, current_row, n_weekly_cols-1, f"{line} Production Plan", line_title_fmt)
            current_row += 1
            
            for col_idx, col_name in enumerate(weekly_output_cols):
                if isinstance(col_name, datetime.date):
                    worksheet.write(current_row, col_idx, col_name, date_head_fmt)
                else:
                    worksheet.write(current_row, col_idx, str(col_name).replace('.',''), header_fmt)
            current_row += 1
            
            sr_no_counter = 1
            for _, row_data in line_df.iterrows():
                for col_idx, col_name in enumerate(weekly_output_cols):
                    mapped_col = "Sr. No." if col_name == "Sr No" else col_name
                    val = sr_no_counter if mapped_col == "Sr. No." else row_data[mapped_col]
                    cell_fmt = data_left_fmt if col_idx in [1, 2] else data_fmt
                    if pd.isna(val):
                        worksheet.write(current_row, col_idx, "", cell_fmt)
                    else:
                        worksheet.write(current_row, col_idx, val, cell_fmt)
                current_row += 1
                sr_no_counter += 1
            
            # Plan Totals
            worksheet.merge_range(current_row, 0, current_row, 2, "Total", total_fmt)
            
            for col_idx in range(3, n_weekly_cols):
                mapped_col = "Sr. No." if weekly_output_cols[col_idx] == "Sr No" else weekly_output_cols[col_idx]
                total_val = line_df[mapped_col].sum()
                worksheet.write(current_row, col_idx, total_val, total_fmt)
            
            current_row += 2 # Empty row space
            
            # --- ACTUAL SECTION ---
            worksheet.merge_range(current_row, 0, current_row, n_weekly_cols-1, f"{line} Production Actual", line_title_fmt)
            current_row += 1
            
            # Headers
            for col_idx, col_name in enumerate(weekly_output_cols):
                if isinstance(col_name, datetime.date):
                    worksheet.write(current_row, col_idx, col_name, date_head_fmt)
                else:
                    worksheet.write(current_row, col_idx, str(col_name).replace('.',''), header_fmt)
            current_row += 1
            
            sr_no_counter = 1
            for _, row_data in line_df.iterrows():
                for col_idx, col_name in enumerate(weekly_output_cols):
                    if isinstance(col_name, datetime.date):
                        # Blank for actuals
                        worksheet.write(current_row, col_idx, "", actual_data_fmt)
                    else:
                        mapped_col = "Sr. No." if col_name == "Sr No" else col_name
                        val = sr_no_counter if mapped_col == "Sr. No." else row_data[mapped_col]
                        cell_fmt = data_left_fmt if col_idx in [1, 2] else actual_data_fmt
                        if pd.isna(val):
                            worksheet.write(current_row, col_idx, "", cell_fmt)
                        else:
                            worksheet.write(current_row, col_idx, val, cell_fmt)
                current_row += 1
                sr_no_counter += 1
            
            # Actual Totals
            worksheet.merge_range(current_row, 0, current_row, 2, "Total", total_fmt)
            worksheet.write(current_row, 3, line_df["Weekly Plan Qty"].sum(), total_fmt)
            for col_idx in range(4, n_weekly_cols):
                worksheet.write(current_row, col_idx, "", total_fmt) # empty yellow cells
            
            current_row += 2 # Empty row space
            
            # --- DOWNTIME SECTION ---
            worksheet.merge_range(current_row, 0, current_row, n_weekly_cols-1, "Major Downtime", downtime_title_fmt)
            current_row += 1
            
            for i in range(1, 4): 
                worksheet.write(current_row, 0, i, data_fmt)
                for col_idx in range(1, n_weekly_cols):
                    worksheet.write(current_row, col_idx, "", data_fmt)
                current_row += 1
            
            current_row += 1 # 1 row margin
            
            # Mark page break after this line
            h_pagebreaks.append(current_row)
        
        # Apply the explicit horizontal page breaks so each line gets its own page segment
        worksheet.set_h_pagebreaks(h_pagebreaks[:-1]) # Omit the last one
            
    return output.getvalue()


def generate_weekly_pdf_report(result_df, final_cols, start_date, selected_week):
    """Generates a linewise PDF report using ReportLab, fitting 1 line per A4 Landscape page."""
    if not REPORTLAB_AVAILABLE:
        return None
        
    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        leftMargin=15,
        rightMargin=15,
        topMargin=20,
        bottomMargin=20
    )
    
    # 842 pt wide total A4 landscape. 842 - 30 margin = 812 usable points
    col_widths = [25, 120, 240, 70] + [58] * 6 # 25+120+240+70+348 = 803 pts. Perfect fit.
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('MainTitle', parent=styles['Normal'], alignment=1, fontSize=12, fontName='Helvetica-Bold')
    company_style = ParagraphStyle('CompanyTitle', parent=styles['Normal'], alignment=1, fontSize=14, fontName='Helvetica')
    format_no_style = ParagraphStyle('FormatNo', parent=styles['Normal'], alignment=1, fontSize=10, fontName='Helvetica')
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], alignment=1, fontSize=8, fontName='Helvetica-Bold', textColor=colors.white)
    data_style = ParagraphStyle('DataStyle', parent=styles['Normal'], alignment=0, fontSize=8, fontName='Helvetica')
    data_center_style = ParagraphStyle('DataCenter', parent=styles['Normal'], alignment=1, fontSize=8, fontName='Helvetica')
    
    date_cols = [c for c in final_cols if isinstance(c, datetime.date)]
    display_cols = ["Sr No", "Part Number", "Part Description", "Weekly Plan Qty"]
    weekly_output_cols = display_cols + date_cols
    
    month_title_txt = f"Weekly Production Plan {start_date.strftime('%B %Y')}"
    
    story = []
    
    unique_lines = result_df["Line Name"].unique()
    
    # Defines the visual style for our tables
    def get_table_style(has_total=True):
        ts = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E78')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ])
        if has_total:
            ts.add('BACKGROUND', (0,-1), (-1,-1), colors.yellow)
            ts.add('TEXTCOLOR', (0,-1), (-1,-1), colors.black)
            ts.add('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold')
            ts.add('SPAN', (0,-1), (2,-1))
        return ts

    for line_idx, line in enumerate(unique_lines):
        line_df = result_df[result_df["Line Name"] == line]
        
        col1_w = sum(col_widths[:-2])
        col2_w = sum(col_widths[-2:])
        
        title_data = [
            [Paragraph("KSPG Automotive India Pvt Ltd , BU Bearing", company_style), Paragraph("Format No:- PRH.F.46.00", format_no_style)],
            [Paragraph(f"<b>{month_title_txt}</b>", title_style), Paragraph(f"<b>{selected_week}</b>", title_style)],
            [Paragraph(f"<b>{line} Production Plan</b>", title_style), ""]
        ]
        
        title_table = Table(title_data, colWidths=[col1_w, col2_w], style=TableStyle([
            ('BOX', (0,0), (-1,-1), 0.5, colors.black),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('SPAN', (0,2), (1,2)) 
        ]))
        story.append(title_table)
        
        headers = [c.strftime('%d-%b') if isinstance(c, datetime.date) else c for c in weekly_output_cols]
        plan_data = [[Paragraph(h, header_style) for h in headers]]
        
        sr_no = 1
        for _, row_data in line_df.iterrows():
            row_items = []
            for col_idx, col_name in enumerate(weekly_output_cols):
                mapped_col = "Sr. No." if col_name == "Sr No" else col_name
                val = sr_no if mapped_col == "Sr. No." else row_data[mapped_col]
                if pd.isna(val): val = ""
                # strings vs numbers
                p_style = data_style if col_idx in [1, 2] else data_center_style
                row_items.append(Paragraph(str(val), p_style))
            plan_data.append(row_items)
            sr_no += 1
            
        # Total Row
        total_row = [Paragraph("<b>Total</b>", data_center_style), Paragraph("", data_center_style), Paragraph("", data_style)]
        for col_idx in range(3, len(weekly_output_cols)):
            mapped_col = "Sr. No." if weekly_output_cols[col_idx] == "Sr No" else weekly_output_cols[col_idx]
            val = line_df[mapped_col].sum()
            total_row.append(Paragraph(f"<b>{val}</b>", data_center_style))
        plan_data.append(total_row)
        
        # Draw plan table
        t = Table(plan_data, colWidths=col_widths, repeatRows=1)
        t.setStyle(get_table_style(has_total=True))
        story.append(t)
        story.append(Spacer(1, 10))
        
        # 3. Actual Section
        story.append(Table([[Paragraph(f"<b>{line} Production Actual</b>", title_style)]], colWidths=[sum(col_widths)], style=TableStyle([
            ('BOX', (0,0), (-1,-1), 0.5, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ])))
        
        actual_data = [[Paragraph(h, header_style) for h in headers]]
        sr_no = 1
        for _, row_data in line_df.iterrows():
            row_items = []
            for col_idx, col_name in enumerate(weekly_output_cols):
                if isinstance(col_name, datetime.date):
                    row_items.append(Paragraph("", data_center_style))
                else:
                    mapped_col = "Sr. No." if col_name == "Sr No" else col_name
                    val = sr_no if mapped_col == "Sr. No." else row_data[mapped_col]
                    if pd.isna(val): val = ""
                    p_style = data_style if col_idx in [1, 2] else data_center_style
                    row_items.append(Paragraph(str(val), p_style))
            actual_data.append(row_items)
            sr_no += 1
            
        actual_total_row = [Paragraph("<b>Total</b>", data_center_style), Paragraph("", data_center_style), Paragraph("", data_style)]
        actual_total_row.append(Paragraph(f"<b>{line_df['Weekly Plan Qty'].sum()}</b>", data_center_style))
        for col_idx in range(4, len(weekly_output_cols)):
            actual_total_row.append(Paragraph("", data_center_style))
        actual_data.append(actual_total_row)
        
        t_actual = Table(actual_data, colWidths=col_widths, repeatRows=1)
        t_actual.setStyle(get_table_style(has_total=True))
        story.append(t_actual)
        story.append(Spacer(1, 10))
        
        # 4. Downtime Section
        story.append(Table([[Paragraph("<b>Major Downtime</b>", title_style)]], colWidths=[sum(col_widths)], style=TableStyle([
            ('BOX', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#1F4E78')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ])))
        
        dt_data = []
        for i in range(1, 4):
            dt_row = [Paragraph(str(i), data_center_style)]
            for _ in range(1, len(weekly_output_cols)):
                dt_row.append(Paragraph("", data_center_style))
            dt_data.append(dt_row)
        
        t_dt = Table(dt_data, colWidths=col_widths, rowHeights=[15, 15, 15])
        t_dt.setStyle(get_table_style(has_total=False))
        story.append(t_dt)
        
        # PAGE BREAK AFTER EVERY LINE
        if line_idx < len(unique_lines) - 1:
            story.append(PageBreak())

    try:
        doc.build(story)
        return output.getvalue()
    except Exception as e:
        print(f"Error building PDF: {e}")
        return None

def render_monthly_plan():
    st.subheader("Monthly Planning")
    START_DATE = datetime.date(2026, 1, 1)
    END_DATE = datetime.date(2026, 1, 31)
    working_days = get_monthly_working_days(START_DATE, END_DATE)

    col1, col2 = st.columns([2, 1])
    with col1:
        st.write("Download the required Excel format to input your production plan.")
    with col2:
        st.download_button(
            label="📄 Download Template (.xlsx)",
            data=get_template_buffer().getvalue(),
            file_name="Production_Plan_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    st.divider()

    uploaded_file = st.file_uploader("Upload filled Excel File (Monthly)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.dataframe(df.head())
        if st.button("Process Monthly Schedule"):
            with st.spinner("Scheduling Production..."):
                result_df = process_schedule(df.copy(), working_days)
                if result_df is not None:
                    meta_cols = ["Sr. No.", "Line Name", "Part Number", "Part Description", "Total Plan Qty", "Major Setup", "Minor Setup"]
                    available_meta = [c for c in meta_cols if c in result_df.columns]
                    date_cols_in_df = [c for c in result_df.columns if isinstance(c, datetime.date)]
                    date_cols_in_df.sort()
                    final_cols = available_meta + date_cols_in_df
                    
                    st.success("Scheduling Complete!")
                    st.dataframe(result_df[final_cols].head(), use_container_width=True)
                    
                    processed_data = generate_monthly_excel(result_df, final_cols, "Datewise Production Plan Jan 2026")
                    st.download_button(
                        label="📥 Download Results (.xlsx)",
                        data=processed_data,
                        file_name="Rheinmetall_Monthly_Schedule.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

def render_weekly_plan():
    st.subheader("Weekly Planning")
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Select Start Date", value=datetime.date.today())
    with col2:
        selected_week = st.selectbox("Select Week", ["Week-1", "Week-2", "Week-3", "Week-4", "Week-5"])
    working_days = get_weekly_working_days(start_date, num_days=6)
    
    st.info(f"Planning Horizon: **{working_days[0].strftime('%A, %d %b %Y')}** to **{working_days[-1].strftime('%A, %d %b %Y')}** (6 working days)")

    col1, col2 = st.columns([2, 1])
    with col1:
        st.write("Download the required Excel format to input your weekly production plan.")
    with col2:
        st.download_button(
            label="📄 Download Template (.xlsx)",
            data=get_template_buffer().getvalue(),
            file_name="Weekly_Plan_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    st.divider()

    uploaded_file = st.file_uploader("Upload filled Excel File (Weekly)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.dataframe(df.head())
        if st.button("Process Weekly Schedule"):
            with st.spinner("Scheduling Production..."):
                result_df = process_schedule(df.copy(), working_days)
                if result_df is not None:
                    date_cols_in_df = [c for c in result_df.columns if isinstance(c, datetime.date)]
                    date_cols_in_df.sort()
                    
                    for d in date_cols_in_df:
                        result_df[d] = result_df[d].round().fillna(0).astype(int)
                    
                    result_df["Weekly Plan Qty"] = result_df[date_cols_in_df].sum(axis=1).astype(int)
                    result_df = result_df[result_df["Weekly Plan Qty"] > 0].reset_index(drop=True)

                    meta_cols = ["Sr. No.", "Line Name", "Part Number", "Part Description", "Weekly Plan Qty", "Major Setup", "Minor Setup"]
                    available_meta = [c for c in meta_cols if c in result_df.columns]
                    final_cols = available_meta + date_cols_in_df
                    
                    st.success("Weekly Scheduling Complete!")
                    
                    display_cols = ["Sr. No.", "Part Number", "Part Description", "Weekly Plan Qty"] + date_cols_in_df
                    st.dataframe(result_df[display_cols].head(), use_container_width=True)
                    
                    processed_excel = generate_weekly_excel_report(result_df, final_cols, start_date, selected_week)
                    processed_pdf = generate_weekly_pdf_report(result_df, final_cols, start_date, selected_week)
                    
                    colexcel, colpdf = st.columns(2)
                    with colexcel:
                        st.download_button(
                            label="📥 Download Weekly Plan (.xlsx)",
                            data=processed_excel,
                            file_name=f"Weekly_Plan_{start_date.strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with colpdf:
                        if processed_pdf:
                            st.download_button(
                                label="📥 Download Linewise Report (.pdf)",
                                data=processed_pdf,
                                file_name=f"Linewise_Weekly_Plan_{start_date.strftime('%Y%m%d')}.pdf",
                                mime="application/pdf"
                            )
                        else:
                            st.error("ReportLab library missing for PDF generation.")

def run_app():
    st.set_page_config(page_title="Datewise Production Planning", layout="wide")

    st.title("🏭 Datewise Production Planning")
    
    with st.sidebar:
        st.header("Navigation")
        mode = st.radio("Select Module:", ["Weekly Plan", "Monthly Plan"])
        
        st.divider()
        st.header("Line Capacities (Units/Day)")
        items = list(CAPACITIES.items())
        col_c1, col_c2 = st.columns(2)
        for i, (line, cap) in enumerate(items):
            if i % 2 == 0:
                col_c1.metric(line, f"{cap:,}")
            else:
                col_c2.metric(line, f"{cap:,}")
        
        st.caption("Operating Time: 1320 mins/day")
        st.caption("Major Setup: 240 mins | Minor Setup: 60 mins")

    if mode == "Monthly Plan":
        render_monthly_plan()
    elif mode == "Weekly Plan":
        render_weekly_plan()

if __name__ == "__main__":
    run_app()

