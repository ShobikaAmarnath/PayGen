# pdf_generator.py
import os
import sys
from pathlib import Path
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

# --- NEW: Helper function to handle paths for PyInstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def register_fonts():
    """Registers the necessary fonts for ReportLab."""
    # Assumes fonts are in a 'fonts' subfolder
    pdfmetrics.registerFont(TTFont("DejaVuSans", resource_path("fonts/DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", resource_path("fonts/DejaVuSans-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVuSans-Oblique", resource_path("fonts/DejaVuSans-Oblique.ttf")))

def create_payslip(employee_data, salary_details):
    """Generates and saves a single PDF payslip."""
    pay_period = employee_data['Period']
    output_dir_name = f"Payslips_{pay_period.strftime('%b_%Y')}" # e.g., Payslips_Apr_2025
    output_dir = Path(output_dir_name)
    output_dir.mkdir(exist_ok=True)

    pdf_path = output_dir / f"{employee_data['Employee_Name']}.pdf"    
    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    elements = []
 
    PAGE_WIDTH, _ = A4
    AVAILABLE_WIDTH = PAGE_WIDTH - 80
    
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle("normal", parent=styles["Normal"], fontName="DejaVuSans", fontSize=9)
    bold_style = ParagraphStyle("bold", parent=normal_style, fontName="DejaVuSans-Bold", alignment=2)
    footer_header_style = ParagraphStyle("bold", parent=normal_style, fontName="DejaVuSans-Bold", fontSize=10)
    in_words_style = ParagraphStyle("in_words", parent=normal_style, fontName="DejaVuSans-Oblique", alignment=2)
    right_align_style = ParagraphStyle("right_align", parent=normal_style, alignment=2) # 2 for RIGHT

    # --- Header and Logo ---
    try:
        logo = Image(resource_path("company_logo.png"), width=250, height=80)
        elements.append(logo)
    except:
        elements.append(Paragraph("<b>LOGO</b>", bold_style))
    
    elements.append(Paragraph(f"<b>{config.COMPANY_NAME}</b>", ParagraphStyle("title", fontName="DejaVuSans-Bold", fontSize=16, alignment=1)))
    elements.append(Spacer(1, 15))
    elements.append(Paragraph(f"<b>Payslip for {employee_data['Period'].strftime('%B %Y')}</b>", ParagraphStyle("heading", fontName="DejaVuSans-Bold", fontSize=13, alignment=1)))
    elements.append(Spacer(1, 40))
    
    # --- Employee Details ---
    details_data = [
        # Left Side Labels     Colon  Values                                            # Right Side Labels    Colon  Values
        ['Employee Name',      ':', employee_data['Employee_Name'],                       'Bank Name',         ':', employee_data['Bank_Name']],
        ['Employee Number',    ':', employee_data['Employee_ID'],                         'Bank Account No',   ':', employee_data['Bank_Account_No']],
        ['Department',         ':', employee_data['Department'],                          'PAN Number',        ':', employee_data['PAN_Number']],
        ['Designation',        ':', employee_data['Designation'],                         'PF Account Number', ':', employee_data['PF_Account_Number']],
        ['Location',           ':', employee_data['Location'],                            'ESI Number',        ':', employee_data['ESI_Number']],
        ['Date of Joining',    ':', employee_data['Date_of_Joining'].strftime('%B %Y'),   'UAN Number',        ':', employee_data['UAN_Number']],
        ['Days Worked',        ':', employee_data['Days_Worked'],                         'LOP Days',          ':', employee_data['LOP_Days']],
    ]
    
    # Define column widths: Label, Colon, Value, Label, Colon, Value
    details_table = Table(details_data, colWidths=[100, 8, 150, 110, 8, 150], hAlign='CENTER', spaceBefore=10)
    details_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'DejaVuSans'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        # Bold the label columns
        ('FONTNAME', (0,0), (0,-1), 'DejaVuSans-Bold'), 
        ('FONTNAME', (3,0), (3,-1), 'DejaVuSans-Bold'),
        # Center align the colon columns
        ('ALIGN', (1,0), (1,-1), 'CENTER'),
        ('ALIGN', (4,0), (4,-1), 'CENTER'),
    ]))

    elements.append(details_table)
    elements.append(Spacer(1, 18))

    # --- Salary Table ---
    data = [
        [Paragraph("<b>Earnings</b>", bold_style), Paragraph("<b>Amount (₹)</b>", bold_style), Paragraph("<b>Deductions</b>", bold_style), Paragraph("<b>Amount (₹)</b>", bold_style)],
        [Paragraph("Basic", right_align_style), Paragraph(f"₹ {salary_details['basic']:,.2f}", right_align_style), Paragraph("EPF", right_align_style), Paragraph(f"₹ {salary_details['epf']:,.2f}", right_align_style)],
        [Paragraph("HRA", right_align_style), Paragraph(f"₹ {salary_details['hra']:,.2f}", right_align_style), Paragraph("Professional Tax", right_align_style), Paragraph(f"₹ {salary_details['prof_tax']:,.2f}", right_align_style)],
        [Paragraph("Special Allowance", right_align_style), Paragraph(f"₹ {salary_details['special_allowance']:,.2f}", right_align_style), Paragraph("Income Tax", right_align_style), Paragraph(f"₹ {salary_details['income_tax']:,.2f}", right_align_style)],
        [Paragraph("<b>Gross Pay</b>", bold_style), Paragraph(f"<b>₹ {salary_details['gross']:,.2f}</b>", bold_style), Paragraph("<b>Total Deductions</b>", bold_style), Paragraph(f"<b>₹ {salary_details['total_ded']:,.2f}</b>", bold_style)],
    ]

    # ... Table styling ...
    table = Table(data, colWidths=[150, 120, 150, 120])
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
    
    summary_data = [
        # Row 1: Net Pay. The first two cells are empty to align with the columns above.
        ['', '', Paragraph("<b>Net Pay</b>", bold_style), Paragraph(f"<b>₹ {salary_details['net']:,.2f}</b>", right_align_style)],
        
        # Row 2: Amount in Words. The text is in the first cell, and we will span it across all columns.
        [Paragraph(f"<b>In Words: </b>{salary_details['net_in_words']}", in_words_style), '', '', '']
    ]

    # Use the same total width and column structure as the main salary table
    summary_table = Table(summary_data, colWidths=[150, 120, 150, 120])
    
    summary_table.setStyle(TableStyle([
        # Style for the Net Pay row (row 0)
        ('BACKGROUND', (2,0), (3,0), colors.lightgreen),
        ('GRID', (2,0), (3,0), 0.5, colors.black),
        ('VALIGN', (2,0), (3,0), 'MIDDLE'),
        ('PADDING', (2,0), (3,0), 6),
        ('FONTSIZE', (2,0), (3,0), 12),

        # Style for the "In Words" row (row 1)
        # Span the cell from the first column (0) to the last column (-1)
        ('SPAN', (0,1), (-1,1)),
        # Add some padding above the text to separate it from the Net Pay line
        ('TOPPADDING', (0,1), (-1,1), 8),
    ]))
    
    elements.append(summary_table)
    elements.append(Spacer(1, 160))

    # ... Footer Table ...
    hr_line = Drawing(width=AVAILABLE_WIDTH, height=5)
    hr_line.add(Line(x1=0, y1=2, x2=AVAILABLE_WIDTH, y2=2, strokeColor=colors.black, strokeWidth=0.5))
    elements.append(hr_line)
    elements.append(Spacer(1, 8))
    
    # Re-create footer content using the correct styles in Paragraphs to show icons
    address_content = f"""
        {config.COMPANY_ADDR_LINE1}<br/>
        {config.COMPANY_ADDR_LINE2}<br/>
        {config.COMPANY_ADDR_LINE3}
    """
    
    

    footer_data = [[
        Paragraph(f"<b>{config.COMPANY_NAME}</b>", footer_header_style), ""
    ],[
        Paragraph(address_content, normal_style),
    ]]

    footer_info_table = Table(footer_data, colWidths=['55%', '45%'], hAlign='LEFT')
    footer_info_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('SPAN', (0,0), (1,0)), # Company name spans both columns
    ]))
    elements.append(footer_info_table)

    doc.build(elements)

    return pdf_path, output_dir