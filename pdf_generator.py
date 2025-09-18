# pdf_generator.py
import os
import sys
from pathlib import Path
import pandas as pd
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Line
import config

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def register_fonts():
    """Registers the necessary fonts for ReportLab."""
    pdfmetrics.registerFont(TTFont("DejaVuSans", resource_path("fonts/DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", resource_path("fonts/DejaVuSans-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVuSans-Oblique", resource_path("fonts/DejaVuSans-Oblique.ttf")))

def create_payslip(employee_data, salary_details):
    """Generates and saves a single PDF payslip with the new layout."""
    print(f"Generating payslip for {employee_data['Employee_Name']}...")
    pay_period = employee_data['Period']
    output_dir_name = f"Payslips_{pay_period.strftime('%b_%Y')}"
    output_dir = Path(output_dir_name)
    output_dir.mkdir(exist_ok=True)

    pdf_path = output_dir / f"{employee_data['Employee_Name']}.pdf"    
    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements = []

    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle("normal", parent=styles["Normal"], fontName="DejaVuSans", fontSize=9, leading=12)
    bold_style = ParagraphStyle("bold", parent=normal_style, fontName="DejaVuSans-Bold")
    right_align_style = ParagraphStyle("right_align", parent=normal_style, alignment=2)
    bold_style_right = ParagraphStyle("bold_right", parent=bold_style, alignment=2)

    # --- NEW: Header table with Logo on left and Company Address on right ---
    try:
        logo = Image(resource_path("company_logo.png"), width=200, height=80)
    except:
        logo = Paragraph("<b>LOGO</b>", bold_style)

    company_address = f"""
        <b>{config.COMPANY_NAME}</b><br/>
        {config.COMPANY_ADDR_LINE1}<br/>
        {config.COMPANY_ADDR_LINE2}<br/>
        {config.COMPANY_ADDR_LINE3}
    """
    header_data = [[logo, Paragraph(company_address, right_align_style)]]
    
    header_table = Table(header_data, colWidths=['50%', '50%'])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (0,0), 'TOP'),
    ]))
    elements.append(header_table)

    elements.append(Paragraph(f"<b>Payslip for {employee_data['Period'].strftime('%B %Y')}</b>", ParagraphStyle("heading", fontName="DejaVuSans-Bold", fontSize=13, alignment=1, spaceBefore=15, spaceAfter=15)))

    # --- Bordered table for Employee Details ---
    doj = employee_data['Date_of_Joining']
    if isinstance(doj, pd.Timestamp):
        doj_str = doj.strftime('%d-%b-%Y')
    else:
        doj_str = str(doj)

    details_data = [
        [Paragraph('<b>Employee Name</b>', bold_style), employee_data['Employee_Name'], Paragraph('<b>Bank Name</b>', bold_style), employee_data['Bank_Name']],
        [Paragraph('<b>Employee Number</b>', bold_style), employee_data['Employee_ID'], Paragraph('<b>Bank Account No</b>', bold_style), employee_data['Bank_Account_No']],
        [Paragraph('<b>Department</b>', bold_style), employee_data['Department'], Paragraph('<b>PAN Number</b>', bold_style), employee_data['PAN_Number']],
        [Paragraph('<b>Designation</b>', bold_style), employee_data['Designation'], Paragraph('<b>PF Account Number</b>', bold_style), employee_data['PF_Account_Number']],
        [Paragraph('<b>Location</b>', bold_style), employee_data['Location'], Paragraph('<b>ESI Number</b>', bold_style), employee_data['ESI_Number']],
        [Paragraph('<b>Date of Joining</b>', bold_style), doj_str, Paragraph('<b>UAN Number</b>', bold_style), employee_data['UAN_Number']],
        [Paragraph('<b>Days Worked</b>', bold_style), employee_data['Days_Worked'], Paragraph('<b>LOP Days</b>', bold_style), employee_data['LOP_Days']],
    ]
    
    details_table = Table(details_data, colWidths=['25%', '25%', '25%', '25%'])
    details_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'DejaVuSans'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        # Style for the label columns (0 and 2)
        ('BACKGROUND', (0,0), (0,-1), colors.lightgrey),
        ('BACKGROUND', (2,0), (2,-1), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        # Style for the value columns (1 and 3)
        ('ALIGN', (1,0), (1,-1), 'LEFT'),
        ('ALIGN', (3,0), (3,-1), 'LEFT'),
    ]))
    elements.append(details_table)
    elements.append(Spacer(1, 20))

    # --- Salary Table ---
    data = [
        [Paragraph("<b>Earnings</b>", bold_style_right), Paragraph("<b>Amount (₹)</b>", bold_style_right), Paragraph("<b>Deductions</b>", bold_style_right), Paragraph("<b>Amount (₹)</b>", bold_style_right)], 
        [Paragraph("Basic", right_align_style), Paragraph(f"₹ {salary_details['basic']:,.2f}", right_align_style), Paragraph("EPF", right_align_style), Paragraph(f"₹ {salary_details['epf']:,.2f}", right_align_style)], 
        [Paragraph("HRA", right_align_style), Paragraph(f"₹ {salary_details['hra']:,.2f}", right_align_style), Paragraph("Professional Tax", right_align_style), Paragraph(f"₹ {salary_details['prof_tax']:,.2f}", right_align_style)], 
        [Paragraph("Special Allowance", right_align_style), Paragraph(f"₹ {salary_details['special_allowance']:,.2f}", right_align_style), Paragraph("Income Tax", right_align_style), Paragraph(f"₹ {salary_details['income_tax']:,.2f}", right_align_style)], 
        [Paragraph("<b>Gross Pay</b>", bold_style_right), Paragraph(f"<b>₹ {salary_details['gross']:,.2f}</b>", bold_style_right), Paragraph("<b>Total Deductions</b>", bold_style_right), Paragraph(f"<b>₹ {salary_details['total_ded']:,.2f}</b>", bold_style_right)], 
    ] 
    
    # ... Table styling ... 
    table = Table(data, colWidths=['25%', '25%', '25%', '25%'])
    table.setStyle(TableStyle([ 
        ("FONTNAME", (0,1), (0,-2), "DejaVuSans"),
        ("FONTNAME", (2,1), (2,-2), "DejaVuSans"),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.whitesmoke), 
    ])) 

    elements.append(table) 
    elements.append(Spacer(1, 20))

    # --- Combined Net Pay and "In Words" table ---
    net_pay_in_words = f"<b>In Words: </b> {salary_details['net_in_words']}"
    summary_data = [[
        Paragraph(f"<b>Net Pay: ₹ {salary_details['net']:,.2f}</b>", bold_style_right),
        Paragraph(net_pay_in_words, right_align_style)
    ]]

    summary_table = Table(summary_data, colWidths=['40%', '60%'])
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.lightgreen),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('FONTNAME', (0,0), (0,0), 'DejaVuSans-Bold'),
    ]))
    elements.append(summary_table)

    doc.build(elements)
    print(f"Successfully created payslip for {employee_data['Employee_Name']}: {pdf_path}")

    return pdf_path, output_dir