import pandas as pd
import io
import calendar
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from fastapi.responses import StreamingResponse
import zipfile

# --- HELPER FUNCTIONS ---
def get_auto_da(m, y):
    if y <= 2020: return 17.0
    elif y == 2021: return 17.0 if m < 7 else 28.0
    elif y == 2022: return 34.0 if m < 7 else 38.0
    elif y == 2023: return 42.0 if m < 7 else 46.0
    elif y == 2024: return 50.0 if m < 7 else 53.0
    elif y == 2025: return 56.0 if m < 7 else 58.0
    elif y == 2026: return 60.0 if m < 7 else 63.0
    return 0.0

def number_to_words(n):
    if n == 0: return "Zero"
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    def get_words(num):
        if num == 0: return ""
        elif num < 20: return ones[num] + " "
        elif num < 100: return tens[num // 10] + " " + get_words(num % 10)
        else: return ones[num // 100] + " Hundred " + get_words(num % 100)
    words = []
    crore = int(n // 10000000)
    n = n % 10000000
    lakh = int(n // 100000)
    n = n % 100000
    thous = int(n // 1000)
    n = int(n % 1000)
    if crore > 0: words.append(get_words(crore).strip() + " Crore ")
    if lakh > 0: words.append(get_words(lakh).strip() + " Lakh ")
    if thous > 0: words.append(get_words(thous).strip() + " Thousand ")
    if n > 0: words.append(get_words(n).strip())
    return "Rupees " + "".join(words).strip() + " Only"

# --- GENERATE PROFESSIONAL EXCEL ---
def generate_dynamic_excel(ministry, dept, office, period_str, active_months, att_dict, final_bill, logs_df, night_hours_per_duty, month_cols):
    wb = Workbook()
    wb.remove(wb.active)
    font_bold = Font(bold=True)
    font_title = Font(bold=True, size=14)
    align_c = Alignment(horizontal='center', vertical='center')
    align_l = Alignment(horizontal='left', vertical='center')
    fill_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    thin = Side(style='thin')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    ws_admin = wb.create_sheet("Part I - Administration")
    ws_admin.page_setup.orientation = ws_admin.ORIENTATION_LANDSCAPE
    ws_admin.page_setup.paperSize = ws_admin.PAPERSIZE_A4
    ws_admin.sheet_properties.pageSetUpPr.fitToPage = True
    ws_admin.page_setup.fitToWidth = 1
    
    row_idx = 1
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = ministry.upper(); ws_admin[f'A{row_idx}'].font = font_title; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 1
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = f"{dept.upper()}, {office.upper()}"; ws_admin[f'A{row_idx}'].font = font_bold; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 1
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = f"Night Duty Allowance Calculation from {period_str}"; ws_admin[f'A{row_idx}'].font = font_bold; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 2
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = "PART I: MONTH-WISE ATTENDANCE REGISTER"; ws_admin[f'A{row_idx}'].font = font_title; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 2
    
    att_cell_refs = {}
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        num_days = calendar.monthrange(y, m)[1]
        date_cols = [str(i) for i in range(1, num_days + 1)]
        att_df = att_dict[(m, y)]
        if att_df.empty: continue
            
        ws_admin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=2+num_days+1)
        ws_admin.cell(row=row_idx, column=1, value=f"ATTENDANCE REGISTER: {m_name.upper()} {y}").font = font_bold
        ws_admin.cell(row=row_idx, column=1).fill = fill_hdr
        row_idx += 1
        headers = ["Emp No", "Name"] + [f"{int(d):02d}" for d in date_cols] + ["Total"]
        for c_i, h in enumerate(headers, 1):
            c = ws_admin.cell(row=row_idx, column=c_i, value=h)
            c.font = font_bold; c.border = border_all; c.alignment = align_c; c.fill = fill_hdr
        row_idx += 1
        
        att_cell_refs[(m, y)] = {}
        for _, emp_row in att_df.iterrows():
            ws_admin.cell(row=row_idx, column=1, value=emp_row['Emp No']).border = border_all
            ws_admin.cell(row=row_idx, column=1).alignment = align_c
            ws_admin.cell(row=row_idx, column=2, value=emp_row['Name']).border = border_all
            ws_admin.cell(row=row_idx, column=2).alignment = align_l
            
            for d_i, d_col in enumerate(date_cols, 3):
                val = "●" if emp_row[d_col] else ""
                c = ws_admin.cell(row=row_idx, column=d_i, value=val)
                c.border = border_all; c.alignment = align_c; c.font = font_bold
            tot_col_idx = 2 + num_days + 1
            tot_col_letter = get_column_letter(tot_col_idx)
            c_tot = ws_admin.cell(row=row_idx, column=tot_col_idx, value=f'=COUNTIF(C{row_idx}:{get_column_letter(2+num_days)}{row_idx}, "●")')
            c_tot.border = border_all; c_tot.font = font_bold; c_tot.alignment = align_c
            att_cell_refs[(m, y)][emp_row['Emp No']] = f"'Part I - Administration'!{tot_col_letter}{row_idx}"
            row_idx += 1
        row_idx += 2
        
    row_idx += 2
    ws_admin.cell(row=row_idx, column=2, value="_______________________").font = font_bold
    ws_admin.cell(row=row_idx+1, column=2, value="Section In-Charge").font = font_bold
    ws_admin.cell(row=row_idx, column=24, value="_______________________").font = font_bold
    ws_admin.cell(row=row_idx+1, column=24, value="Office Spdt/Estb In-Charge").font = font_bold
        
    ws_admin.column_dimensions['A'].width = 10
    ws_admin.column_dimensions['B'].width = 22
    for i in range(3, 3+31): ws_admin.column_dimensions[get_column_letter(i)].width = 3.2
    ws_admin.column_dimensions[get_column_letter(3+31)].width = 6

    # --- PART II ---
    ws_fin = wb.create_sheet("Part II - Finance")
    ws_fin.page_setup.orientation = ws_fin.ORIENTATION_LANDSCAPE
    ws_fin.page_setup.paperSize = ws_fin.PAPERSIZE_A4
    ws_fin.sheet_properties.pageSetUpPr.fitToPage = True
    ws_fin.page_setup.fitToWidth = 1
    
    row_idx = 1
    ws_fin.merge_cells('A1:I1'); ws_fin['A1'] = ministry.upper(); ws_fin['A1'].font = font_title; ws_fin['A1'].alignment = align_c
    row_idx += 1
    ws_fin.merge_cells('A2:I2'); ws_fin['A2'] = f"{dept.upper()}, {office.upper()}"; ws_fin['A2'].font = font_bold; ws_fin['A2'].alignment = align_c
    row_idx += 2
    ws_fin.merge_cells('A3:I3'); ws_fin['A3'] = "PART II: MONTH-WISE FINANCIAL BILL & SUMMARY"; ws_fin['A3'].font = font_title; ws_fin['A3'].alignment = align_c
    row_idx += 2
    
    fin_refs = {}
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        month_logs = logs_df[(logs_df["Year"] == y) & (logs_df["Month"] == m_name)]
        if month_logs.empty: continue
        da_val = get_auto_da(m, y)
        
        ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=9)
        ws_fin.cell(row=row_idx, column=1, value=f"BILL FOR {m_name.upper()} {y} (DA: {da_val}%)").font = font_bold
        ws_fin.cell(row=row_idx, column=1).fill = fill_hdr
        row_idx += 1
        headers = ["Emp No", "Name", "Post", "Basic Pay", "DA %", "Duties", "1 Hr Rate", f"{int(night_hours_per_duty*10)} min Rate", "Net NDA"]
        for c_i, h in enumerate(headers, 1):
            c = ws_fin.cell(row=row_idx, column=c_i, value=h)
            c.font = font_bold; c.border = border_all; c.alignment = align_c; c.fill = fill_hdr
        row_idx += 1
        
        fin_refs[(m, y)] = {}
        start_sum_row = row_idx
        for _, row in month_logs.iterrows():
            emp_no = row['Emp No']
            ws_fin.cell(row=row_idx, column=1, value=emp_no).border = border_all; ws_fin.cell(row=row_idx, column=1).alignment = align_c
            ws_fin.cell(row=row_idx, column=2, value=row['Name']).border = border_all; ws_fin.cell(row=row_idx, column=2).alignment = align_l
            ws_fin.cell(row=row_idx, column=3, value=row['Post']).border = border_all; ws_fin.cell(row=row_idx, column=3).alignment = align_l
            
            c_bp = ws_fin.cell(row=row_idx, column=4, value=row['Basic Pay'])
            c_bp.border = border_all; c_bp.alignment = align_c; c_bp.number_format = '#,##0'
            
            c_da = ws_fin.cell(row=row_idx, column=5, value=da_val)
            c_da.border = border_all; c_da.alignment = align_c
            
            duty_ref = att_cell_refs.get((m, y), {}).get(emp_no, "0")
            c_duty = ws_fin.cell(row=row_idx, column=6, value=f"={duty_ref}")
            c_duty.border = border_all; c_duty.font = font_bold; c_duty.alignment = align_c
            
            c_1hr = ws_fin.cell(row=row_idx, column=7, value=f"=ROUND((MIN(D{row_idx}, 43600)*(1+E{row_idx}/100))/200, 2)")
            c_1hr.border = border_all; c_1hr.alignment = align_c; c_1hr.number_format = '0.00'
            
            c_shift = ws_fin.cell(row=row_idx, column=8, value=f"=ROUND(G{row_idx}*({night_hours_per_duty}/6), 2)")
            c_shift.border = border_all; c_shift.alignment = align_c; c_shift.number_format = '0.00'
            
            c_net = ws_fin.cell(row=row_idx, column=9, value=f"=ROUND(H{row_idx}*F{row_idx}, 2)")
            c_net.border = border_all; c_net.font = font_bold; c_net.alignment = align_c; c_net.number_format = '#,##0.00'
            
            fin_refs[(m, y)][emp_no] = {"duty": f"F{row_idx}", "net": f"I{row_idx}"}
            row_idx += 1
            
        ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)
        ws_fin.cell(row=row_idx, column=1, value=f"TOTAL NDA FOR {m_name.upper()} {y}:").font = font_bold
        ws_fin.cell(row=row_idx, column=1).alignment = Alignment(horizontal='right')
        c_month_sum = ws_fin.cell(row=row_idx, column=9, value=f"=SUM(I{start_sum_row}:I{row_idx-1})")
        c_month_sum.font = font_bold; c_month_sum.alignment = align_c; c_month_sum.number_format = '#,##0.00'
        row_idx += 3

    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=9)
    ws_fin.cell(row=row_idx, column=1, value="GRAND FINAL SUMMARY").font = font_bold
    ws_fin.cell(row=row_idx, column=1).fill = fill_hdr
    row_idx += 1
    sum_headers = ["Emp No", "Name", "Post"] + month_cols + ["Total Duties", "Total Hrs", "Wt Hrs", "Total Net NDA"]
    for c_i, h in enumerate(sum_headers, 1):
        c = ws_fin.cell(row=row_idx, column=c_i, value=h)
        c.font = font_bold; c.border = border_all; c.alignment = align_c; c.fill = fill_hdr
    row_idx += 1
    
    start_sum_row = row_idx
    grand_total_all_actual = 0.0 
    for _, row in final_bill.iterrows():
        emp_no = row['Emp No']
        ws_fin.cell(row=row_idx, column=1, value=emp_no).border = border_all; ws_fin.cell(row=row_idx, column=1).alignment = align_c
        ws_fin.cell(row=row_idx, column=2, value=row['Name']).border = border_all; ws_fin.cell(row=row_idx, column=2).alignment = align_l
        ws_fin.cell(row=row_idx, column=3, value=row['Post']).border = border_all; ws_fin.cell(row=row_idx, column=3).alignment = align_l
        
        c_idx = 4
        duty_cells, net_cells = [], []
        for period in active_months:
            m, y = period['m'], period['y']
            ref = fin_refs.get((m, y), {}).get(emp_no)
            c = ws_fin.cell(row=row_idx, column=c_idx)
            c.border = border_all; c.alignment = align_c
            if ref:
                c.value = f"={ref['duty']}"
                duty_cells.append(get_column_letter(c_idx) + str(row_idx))
                net_cells.append(ref['net'])
            else: c.value = 0
            c_idx += 1
            
        c_tot_d = ws_fin.cell(row=row_idx, column=c_idx, value=f"=SUM({','.join(duty_cells)})" if duty_cells else 0)
        c_tot_d.border = border_all; c_tot_d.font = font_bold; c_tot_d.alignment = align_c
        c_tot_h = ws_fin.cell(row=row_idx, column=c_idx+1, value=f"={get_column_letter(c_idx)}{row_idx}*{night_hours_per_duty}")
        c_tot_h.border = border_all; c_tot_h.alignment = align_c
        c_wt_h = ws_fin.cell(row=row_idx, column=c_idx+2, value=f"={get_column_letter(c_idx+1)}{row_idx}/6")
        c_wt_h.border = border_all; c_wt_h.alignment = align_c; c_wt_h.number_format = '0.00'
        c_net = ws_fin.cell(row=row_idx, column=c_idx+3, value=f"=SUM({','.join(net_cells)})" if net_cells else 0)
        c_net.border = border_all; c_net.font = font_bold; c_net.alignment = align_c; c_net.number_format = '#,##0.00'
        
        grand_total_all_actual += row['Net Payable NDA (₹)']
        row_idx += 1
        
    last_col_l = get_column_letter(c_idx+3)
    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=c_idx+2)
    ws_fin.cell(row=row_idx, column=1, value="GRAND TOTAL PAYOUT:").font = font_bold
    ws_fin.cell(row=row_idx, column=1).alignment = Alignment(horizontal='right')
    c_exact_sum = ws_fin.cell(row=row_idx, column=c_idx+3, value=f'=SUM({last_col_l}{start_sum_row}:{last_col_l}{row_idx-1})')
    c_exact_sum.font = font_bold; c_exact_sum.alignment = align_c; c_exact_sum.number_format = '#,##0.00'
    row_idx += 1

    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=c_idx+2)
    ws_fin.cell(row=row_idx, column=1, value="ROUNDED GRAND TOTAL:").font = font_bold
    ws_fin.cell(row=row_idx, column=1).alignment = Alignment(horizontal='right')
    c_rnd_sum = ws_fin.cell(row=row_idx, column=c_idx+3, value=f'=ROUND({last_col_l}{row_idx-1}, 0)')
    c_rnd_sum.font = font_bold; c_rnd_sum.alignment = align_c; c_rnd_sum.number_format = '#,##0'
    row_idx += 1
    
    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=c_idx+3)
    ws_fin.cell(row=row_idx, column=1, value=f"({number_to_words(round(grand_total_all_actual))})").font = Font(bold=True, italic=True)
    ws_fin.cell(row=row_idx, column=1).alignment = align_c

    ws_fin.column_dimensions['A'].width = 10
    ws_fin.column_dimensions['B'].width = 22
    ws_fin.column_dimensions['C'].width = 10
    ws_fin.column_dimensions['D'].width = 10
    ws_fin.column_dimensions['G'].width = 10
    ws_fin.column_dimensions['H'].width = 12
    ws_fin.column_dimensions['I'].width = 12
    ws_fin.column_dimensions[last_col_l].width = 12

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- GENERATE WORD (DOC) ---
def generate_word_html(ministry, dept, office, period_str, active_months, att_dict, final_bill, logs_df, night_hours_per_duty, month_cols):
    rate_header = f"{int(night_hours_per_duty * 10)} min Rate"
    html = f"""<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head><style>
        body {{ font-family: Arial, sans-serif; font-size: 10pt; }}
        h3, h4, h5 {{ text-align: center; margin: 5px 0; }}
        table {{ border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 25px; border: 1px solid black; }}
        th, td {{ border: 1px solid black; padding: 4px; text-align: center; font-size: 9pt; vertical-align: middle; }}
        th {{ background-color: #d9e1f2; font-weight: bold; }}
        .header-row {{ background-color: #2f5597; color: white; font-weight: bold; padding: 5px; text-align: left; }}
    </style></head><body>
    <h3>{ministry.upper()}</h3><h4>{dept.upper()}, {office.upper()}</h4><h5>Night Duty Allowance Calculation from {period_str}</h5><br>"""
    
    html += "<h4>PART I: MONTH-WISE ATTENDANCE REGISTER</h4>"
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        num_days = calendar.monthrange(y, m)[1]
        date_cols = [str(i) for i in range(1, num_days + 1)]
        att_df = att_dict[(m, y)]
        if att_df.empty: continue
            
        html += f"<div class='header-row'>ATTENDANCE REGISTER: {m_name.upper()} {y}</div>"
        html += "<table><tr><th>Emp No</th><th>Name</th>"
        for d in range(1, len(date_cols) + 1): html += f"<th>{d:02d}</th>"
        html += "<th>Total</th></tr>"
        for _, row in att_df.iterrows():
            duties_count = sum(1 for d_col in date_cols if row[d_col])
            if duties_count == 0: continue
            html += f"<tr><td>{row['Emp No']}</td><td style='text-align:left;'>{row['Name'][:20]}</td>"
            for d_col in date_cols:
                mark = "<span style='font-size:14pt;'>&#9679;</span>" if row[d_col] else ""
                html += f"<td>{mark}</td>"
            html += f"<td style='font-weight:bold;'>{duties_count}</td></tr>"
        html += "</table><br>"

    html += f"<h3 style='page-break-before: always;'>{ministry.upper()}</h3>"
    html += f"<h4>{dept.upper()}, {office.upper()}</h4>"
    html += "<h4>PART II: MONTH-WISE FINANCIAL BILL & GRAND SUMMARY</h4>"
    
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        month_logs = logs_df[(logs_df["Year"] == y) & (logs_df["Month"] == m_name)]
        if month_logs.empty: continue
        da_val = get_auto_da(m, y)
        
        html += f"<div class='header-row'>BILL FOR {m_name.upper()} {y} (DA: {da_val}%)</div>"
        html += f"<table><tr><th>Emp No</th><th>Name</th><th>Post</th><th>Basic Pay</th><th>DA %</th><th>Duties</th><th>1 Hr Rate</th><th>{rate_header}</th><th>Net NDA (&#8377;)</th></tr>"
        month_total_amt = 0.0
        for _, row in month_logs.iterrows():
            bp = row["Basic Pay"]
            hr_rate = (min(bp, 43600) * (1 + da_val/100)) / 200
            day_rate = hr_rate * (night_hours_per_duty / 6)
            net_amt = row["NDA Paid (₹)"]
            month_total_amt += net_amt
            html += f"<tr><td>{row['Emp No']}</td><td style='text-align:left;'>{row['Name']}</td><td style='text-align:left;'>{row['Post']}</td><td>{bp:,.0f}</td><td>{int(da_val)}</td><td style='font-weight:bold;'>{row['Duties']}</td><td>{hr_rate:,.2f}</td><td>{day_rate:,.2f}</td><td style='font-weight:bold;'>{net_amt:,.2f}</td></tr>"
        html += f"<tr><td colspan='8' style='text-align:right; font-weight:bold;'>TOTAL NDA FOR {m_name.upper()} {y}:</td><td style='font-weight:bold;'>&#8377; {month_total_amt:,.2f}</td></tr></table>"

    html += "<br><h4>GRAND FINAL SUMMARY</h4>"
    html += "<table><tr><th>Emp No</th><th>Name</th><th>Post</th>"
    for mc in month_cols: html += f"<th>{mc}</th>"
    html += "<th>Total Duties</th><th>Total Hrs</th><th>Wt Hrs</th><th>Total Net NDA (&#8377;)</th></tr>"
    grand_total_all = 0.0
    for _, row in final_bill.iterrows():
        total_hrs = row['Total Night Duties'] * night_hours_per_duty
        wt_hrs = total_hrs / 6
        html += f"<tr><td>{row['Emp No']}</td><td style='text-align:left;'>{row['Name']}</td><td style='text-align:left;'>{row['Post']}</td>"
        for mc in month_cols: html += f"<td>{row[mc]}</td>"
        html += f"<td style='font-weight:bold;'>{row['Total Night Duties']}</td><td>{total_hrs}</td><td>{wt_hrs:,.2f}</td><td style='font-weight:bold;'>{row['Net Payable NDA (₹)']:,.2f}</td></tr>"
        grand_total_all += row['Net Payable NDA (₹)']
        
    rounded_grand_total = round(grand_total_all)
    amount_words = number_to_words(rounded_grand_total)
    
    html += f"<tr><td colspan='{6 + len(month_cols)}' style='text-align:right; font-weight:bold;'>GRAND TOTAL PAYOUT:</td><td style='font-weight:bold; font-size:11pt;'>&#8377; {grand_total_all:,.2f}</td></tr>"
    html += f"<tr><td colspan='{6 + len(month_cols)}' style='text-align:right; font-weight:bold;'>ROUNDED GRAND TOTAL:</td><td style='font-weight:bold; font-size:11pt;'>&#8377; {rounded_grand_total:,.0f}</td></tr>"
    html += f"<tr><td colspan='{7 + len(month_cols)}' style='text-align:center; font-weight:bold; font-style:italic;'>({amount_words})</td></tr></table>"
    html += "</body></html>"
    return html.encode('utf-8')

# --- GENERATE PDF ---
def generate_official_pdf(ministry, dept, office, period_str, active_months, att_dict, final_bill, logs_df, night_hours_per_duty, month_cols):
    pdf = FPDF(orientation='L', unit='mm', format='A4') 
    
    # --- PART 1: ATTENDANCE ---
    pdf.add_page()
    pdf.set_font("Arial", 'B', 15)
    pdf.cell(0, 7, txt=ministry.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 6, txt=dept.upper() + ", " + office.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'BU', 11)
    pdf.cell(0, 7, txt=f"Night Duty Allowance Calculation from {period_str}", ln=True, align='C')
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, txt="PART I: MONTH-WISE ATTENDANCE REGISTER", ln=True, align='C')
    pdf.ln(3)

    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        num_days = calendar.monthrange(y, m)[1]
        date_cols = [str(i) for i in range(1, num_days + 1)]
        att_df = att_dict[(m, y)]
        if att_df.empty: continue
            
        if pdf.get_y() > 140: pdf.add_page()
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(220, 230, 240) 
        pdf.cell(0, 7, txt=f" ATTENDANCE REGISTER: {m_name.upper()} {y} ", ln=True, align='L', fill=True)
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(240, 240, 240)
        day_w = 6.2
        pdf.cell(14, 6, "Emp No", 1, 0, 'C', fill=True)
        pdf.cell(45, 6, "Name", 1, 0, 'C', fill=True)
        for d in range(1, num_days + 1): pdf.cell(day_w, 6, f"{d:02d}", 1, 0, 'C', fill=True)
        pdf.cell(12, 6, "Total", 1, 1, 'C', fill=True)
        
        for _, row in att_df.iterrows():
            duties_count = sum(1 for d_col in date_cols if row[d_col])
            if duties_count == 0: continue
            pdf.set_font("Arial", '', 8)
            pdf.cell(14, 6, str(row['Emp No']), 1, 0, 'C')
            pdf.cell(45, 6, str(row['Name'])[:28], 1, 0, 'L')
            for d_col in date_cols:
                if row[d_col]:
                    pdf.set_font("Arial", '', 13) 
                    pdf.cell(day_w, 6, chr(149), 1, 0, 'C')
                    pdf.set_font("Arial", '', 8) 
                else:
                    pdf.cell(day_w, 6, "", 1, 0, 'C')
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(12, 6, str(duties_count), 1, 1, 'C')
        pdf.ln(5)

    # --- PART 2: BILL & SUMMARY ---
    pdf.add_page()
    pdf.set_font("Arial", 'B', 15)
    pdf.cell(0, 7, txt=ministry.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 6, txt=dept.upper() + ", " + office.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'BU', 12)
    pdf.cell(0, 8, txt="PART II: MONTH-WISE FINANCIAL BILL & GRAND SUMMARY", ln=True, align='C')
    pdf.ln(3)

    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        month_logs = logs_df[(logs_df["Year"] == y) & (logs_df["Month"] == m_name)]
        if month_logs.empty: continue
        da_val = get_auto_da(m, y)
        
        if pdf.get_y() > 150: pdf.add_page()
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(220, 230, 240)
        pdf.cell(0, 7, txt=f" BILL FOR {m_name.upper()} {y} (DA: {da_val}%) ", ln=True, align='L', fill=True)
        pdf.set_font("Arial", 'B', 9)
        pdf.set_fill_color(240, 240, 240)
        bw = [18, 55, 25, 25, 15, 20, 25, 30, 30] 
        pdf.cell(bw[0], 6, "Emp No", 1, 0, 'C', fill=True)
        pdf.cell(bw[1], 6, "Name", 1, 0, 'C', fill=True)
        pdf.cell(bw[2], 6, "Post", 1, 0, 'C', fill=True)
        pdf.cell(bw[3], 6, "Basic Pay", 1, 0, 'C', fill=True)
        pdf.cell(bw[4], 6, "DA%", 1, 0, 'C', fill=True)
        pdf.cell(bw[5], 6, "Duties", 1, 0, 'C', fill=True)
        pdf.cell(bw[6], 6, "1 Hr Rate", 1, 0, 'C', fill=True)
        pdf.cell(bw[7], 6, f"{int(night_hours_per_duty * 10)} min Rate", 1, 0, 'C', fill=True)
        pdf.cell(bw[8], 6, "Net NDA", 1, 1, 'C', fill=True)
        
        pdf.set_font("Arial", '', 9)
        month_total_amt = 0.0
        for _, row in month_logs.iterrows():
            bp = row["Basic Pay"]
            hr_rate = (min(bp, 43600) * (1 + da_val/100)) / 200
            day_rate = hr_rate * (night_hours_per_duty / 6)
            net_amt = row["NDA Paid (₹)"]
            month_total_amt += net_amt
            pdf.cell(bw[0], 7, str(row['Emp No']), 1, 0, 'C')
            pdf.cell(bw[1], 7, str(row['Name'])[:30], 1, 0, 'L')
            pdf.cell(bw[2], 7, str(row['Post']), 1, 0, 'L')
            pdf.cell(bw[3], 7, f"{bp:,.0f}", 1, 0, 'C')
            pdf.cell(bw[4], 7, str(int(da_val)), 1, 0, 'C')
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(bw[5], 7, str(row['Duties']), 1, 0, 'C')
            pdf.set_font("Arial", '', 9)
            pdf.cell(bw[6], 7, f"{hr_rate:,.2f}", 1, 0, 'C')
            pdf.cell(bw[7], 7, f"{day_rate:,.2f}", 1, 0, 'C')
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(bw[8], 7, f"Rs. {net_amt:,.2f}", 1, 1, 'C')
            pdf.set_font("Arial", '', 9)
            
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(sum(bw[:-1]), 8, f"TOTAL NDA FOR {m_name.upper()} {y}:   ", 1, 0, 'R')
        pdf.cell(bw[-1], 8, f"Rs. {month_total_amt:,.2f}", 1, 1, 'C')
        pdf.ln(5) 

    if pdf.get_y() > 130: pdf.add_page()
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, txt="GRAND FINAL SUMMARY", ln=True, align='C')
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(220, 230, 240)
    sw_base = [15, 45, 15] 
    sw_end = [16, 20, 20, 25] 
    usable_width = 277 - sum(sw_base) - sum(sw_end)
    m_w = usable_width / len(month_cols) if month_cols else 15
    
    pdf.cell(sw_base[0], 7, "Emp No", 1, 0, 'C', fill=True)
    pdf.cell(sw_base[1], 7, "Name", 1, 0, 'C', fill=True)
    pdf.cell(sw_base[2], 7, "Post", 1, 0, 'C', fill=True)
    for mc in month_cols: pdf.cell(m_w, 7, mc, 1, 0, 'C', fill=True)
    pdf.cell(sw_end[0], 7, "Tot Duties", 1, 0, 'C', fill=True)
    pdf.cell(sw_end[1], 7, "Total Hrs", 1, 0, 'C', fill=True)
    pdf.cell(sw_end[2], 7, "Wt Hrs", 1, 0, 'C', fill=True)
    pdf.cell(sw_end[3], 7, "Net NDA", 1, 1, 'C', fill=True)
    
    pdf.set_font("Arial", '', 9)
    grand_total_all = 0.0
    for _, row in final_bill.iterrows():
        total_hrs = row['Total Night Duties'] * night_hours_per_duty
        wt_hrs = total_hrs / 6
        pdf.cell(sw_base[0], 7, str(row['Emp No']), 1, 0, 'C')
        pdf.cell(sw_base[1], 7, row['Name'][:25], 1, 0, 'L')
        pdf.cell(sw_base[2], 7, row['Post'], 1, 0, 'L')
        pdf.set_font("Arial", 'B', 9)
        for mc in month_cols: pdf.cell(m_w, 7, str(row[mc]), 1, 0, 'C')
        pdf.cell(sw_end[0], 7, str(row['Total Night Duties']), 1, 0, 'C')
        pdf.set_font("Arial", '', 9)
        pdf.cell(sw_end[1], 7, str(total_hrs), 1, 0, 'C')
        pdf.cell(sw_end[2], 7, f"{wt_hrs:,.2f}", 1, 0, 'C')
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(sw_end[3], 7, f"Rs. {row['Net Payable NDA (₹)']:,.2f}", 1, 1, 'C')
        pdf.set_font("Arial", '', 9)
        grand_total_all += row['Net Payable NDA (₹)']
        
    pdf.set_font("Arial", 'B', 11)
    total_w_before_amt = sum(sw_base) + (m_w * len(month_cols)) + sum(sw_end[:-1])
    rounded_grand_total = round(grand_total_all)
    amount_words = number_to_words(rounded_grand_total)
    
    pdf.cell(total_w_before_amt, 9, "GRAND TOTAL PAYOUT:   ", 1, 0, 'R')
    pdf.cell(sw_end[-1], 9, f"Rs. {grand_total_all:,.2f}", 1, 1, 'C')
    pdf.cell(total_w_before_amt, 9, "ROUNDED GRAND TOTAL:   ", 1, 0, 'R')
    pdf.cell(sw_end[-1], 9, f"Rs. {rounded_grand_total:,.0f}", 1, 1, 'C')
    pdf.set_font("Arial", 'BI', 10)
    pdf.cell(0, 8, f"({amount_words})", 0, 1, 'C')

    # --- PART 3: EMPLOYEE-WISE FORMAT ---
    pdf.add_page()
    pdf.set_font("Arial", 'B', 15)
    pdf.cell(0, 7, txt=ministry.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 6, txt=dept.upper() + ", " + office.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'BU', 12)
    pdf.cell(0, 8, txt="PART III: EMPLOYEE-WISE FINANCIAL FORMAT", ln=True, align='C')
    pdf.ln(3)

    unique_emps = logs_df['Emp No'].unique()
    for emp in unique_emps:
        emp_logs = logs_df[logs_df['Emp No'] == emp]
        emp_name = emp_logs.iloc[0]['Name']
        emp_post = emp_logs.iloc[0]['Post']
        
        if pdf.get_y() > 150: 
            pdf.add_page()
            
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(240, 240, 240)
        pdf.cell(0, 6, f"NIGHT DUTY ALLOWANCE FOR THE PERIOD {period_str}", 1, 1, 'C', fill=True)
        pdf.cell(0, 6, f"PN- {emp} {str(emp_name).upper()}, {str(emp_post).upper()}", 1, 1, 'C', fill=True)
        
        pdf.set_font("Arial", 'B', 8)
        w3 = [12, 16, 20, 12, 22, 22, 35, 30, 45, 25, 38] 
        
        pdf.cell(w3[0], 10, "Sr. No.", 1, 0, 'C')
        pdf.cell(w3[1], 10, "Period", 1, 0, 'C')
        pdf.cell(w3[2], 10, "Basic Pay", 1, 0, 'C')
        pdf.cell(w3[3], 10, "D.A%", 1, 0, 'C')
        pdf.cell(w3[4], 10, "DA in Rs.", 1, 0, 'C')
        pdf.cell(w3[5], 10, "BP+DA", 1, 0, 'C')
        pdf.cell(w3[6], 10, "(BP+DA)/200 For 1hr", 1, 0, 'C')
        pdf.cell(w3[7], 10, "NDA 1 Minute", 1, 0, 'C')
        
        rate_mins = int(night_hours_per_duty * 10)
        pdf.cell(w3[8], 10, f"NDA for one day ({rate_mins}Min)", 1, 0, 'C')
        pdf.cell(w3[9], 10, "Total Days", 1, 0, 'C')
        pdf.cell(w3[10], 10, "Total Amount", 1, 1, 'C')
        
        pdf.set_font("Arial", '', 8)
        sr_no = 1
        emp_total = 0.0
        emp_days = 0
        for _, r in emp_logs.iterrows():
            bp = r["Basic Pay"]
            capped_bp = min(bp, 43600)
            da_pct = r["DA %"]
            da_rs = capped_bp * (da_pct / 100)
            bp_da = capped_bp + da_rs
            hr_rate = bp_da / 200
            min_rate = hr_rate / 60
            day_rate = hr_rate * (night_hours_per_duty / 6)
            days = r["Duties"]
            amt = r["NDA Paid (₹)"]
            
            month_abbr = r["Month"][:3]
            yr_str = str(r["Year"])[-2:]
            period_val = f"{month_abbr}-{yr_str}"
            
            pdf.cell(w3[0], 7, str(sr_no), 1, 0, 'C')
            pdf.cell(w3[1], 7, period_val, 1, 0, 'C')
            pdf.cell(w3[2], 7, f"{bp:,.0f}", 1, 0, 'C')
            pdf.cell(w3[3], 7, str(int(da_pct)), 1, 0, 'C')
            pdf.cell(w3[4], 7, f"{da_rs:,.0f}", 1, 0, 'C')
            pdf.cell(w3[5], 7, f"{bp_da:,.0f}", 1, 0, 'C')
            pdf.cell(w3[6], 7, f"{hr_rate:,.2f}", 1, 0, 'C')
            pdf.cell(w3[7], 7, f"{min_rate:,.2f}", 1, 0, 'C')
            pdf.cell(w3[8], 7, f"{day_rate:,.2f}", 1, 0, 'C')
            pdf.cell(w3[9], 7, str(days), 1, 0, 'C')
            pdf.cell(w3[10], 7, f"{amt:,.2f}", 1, 1, 'C')
            
            sr_no += 1
            emp_total += amt
            emp_days += days
            
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(sum(w3[:9]), 7, "TOTAL", 1, 0, 'C')
        pdf.cell(w3[9], 7, str(emp_days), 1, 0, 'C')
        pdf.cell(w3[10], 7, f"{emp_total:,.2f}", 1, 1, 'C')
        pdf.ln(8)

    # Note: Returns clean byte output without encode (fixed earlier)
    return pdf.output()

# --- MAIN API HANDLER ---
def process_and_generate_reports(data):
    active_months = [{"m": m_data.month, "y": m_data.year} for m_data in data.months_data]
    period_str = f"{calendar.month_name[active_months[0]['m']]} {active_months[0]['y']} to {calendar.month_name[active_months[-1]['m']]} {active_months[-1]['y']}"
    
    att_dict = {}
    logs_list = []
    month_cols = []
    
    for m_data in data.months_data:
        m, y = m_data.month, m_data.year
        m_name = calendar.month_name[m]
        col_name = f"{calendar.month_abbr[m]} '{str(y)[-2:]}"
        month_cols.append(col_name)
        da_val = get_auto_da(m, y)
        
        att_rows = []
        for rec in m_data.records:
            row = {"Emp No": rec.emp_no, "Name": rec.name, "Post": rec.post, "Basic Pay": rec.basic_pay}
            num_days = calendar.monthrange(y, m)[1]
            for d in range(1, num_days + 1):
                row[str(d)] = rec.attendance.get(str(d), False)
            att_rows.append(row)
            
            if rec.total_duties > 0:
                capped_bp = min(rec.basic_pay, 43600)
                hr_rate = (capped_bp * (1 + da_val/100)) / 200
                day_rate = hr_rate * (data.night_hours_per_duty / 6.0)
                net_amt = day_rate * rec.total_duties
                logs_list.append({
                    "Emp No": rec.emp_no, "Name": rec.name, "Post": rec.post,
                    "Basic Pay": rec.basic_pay, "Month": m_name, "Year": y, 
                    "DA %": da_val, "Duties": rec.total_duties, "NDA Paid (₹)": net_amt
                })
        att_dict[(m, y)] = pd.DataFrame(att_rows)

    if not logs_list:
        return {"error": "No duties marked!"}

    logs_df = pd.DataFrame(logs_list)
    final_bill = logs_df.groupby(["Emp No", "Name", "Post"]).agg(Total_Night_Duties=("Duties", "sum"), Net_Payable_NDA=("NDA Paid (₹)", "sum")).reset_index()
    
    for m_data in data.months_data:
        m, y = m_data.month, m_data.year
        col_name = f"{calendar.month_abbr[m]} '{str(y)[-2:]}"
        month_duties = logs_df[(logs_df["Month"] == calendar.month_name[m]) & (logs_df["Year"] == y)].groupby("Emp No")["Duties"].sum().reset_index()
        month_duties.rename(columns={"Duties": col_name}, inplace=True)
        final_bill = pd.merge(final_bill, month_duties, on="Emp No", how="left").fillna({col_name: 0})
        final_bill[col_name] = final_bill[col_name].astype(int)

    final_col_order = ["Emp No", "Name", "Post"] + month_cols + ["Total_Night_Duties", "Net_Payable_NDA"]
    final_bill = final_bill[final_col_order]
    final_bill.rename(columns={"Total_Night_Duties": "Total Night Duties", "Net_Payable_NDA": "Net Payable NDA (₹)"}, inplace=True)

    # File Generation
    excel_bytes = generate_dynamic_excel(data.ministry, data.department, data.office, period_str, active_months, att_dict, final_bill, logs_df, data.night_hours_per_duty, month_cols)
    pdf_bytes = generate_official_pdf(data.ministry, data.department, data.office, period_str, active_months, att_dict, final_bill, logs_df, data.night_hours_per_duty, month_cols)
    word_bytes = generate_word_html(data.ministry, data.department, data.office, period_str, active_months, att_dict, final_bill, logs_df, data.night_hours_per_duty, month_cols)

    # Pack into ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(f"CleanReport_NDA_{period_str}.xlsx", excel_bytes)
        zip_file.writestr(f"CleanReport_NDA_{period_str}.pdf", pdf_bytes)
        zip_file.writestr(f"CleanReport_NDA_{period_str}.doc", word_bytes)
    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer, 
        media_type="application/zip", 
        headers={"Content-Disposition": f"attachment; filename=NDA_Reports_{period_str}.zip"}
    )