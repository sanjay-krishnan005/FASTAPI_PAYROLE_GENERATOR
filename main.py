import pandas as pd
import smtplib
import ssl
import io
import os
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ReportLab Imports
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# FastAPI Imports
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
import uvicorn

MOCK_SECRETS = {
    "smtp": {
        "email": "assassinnightcap008@gmail.com",
        "password": "poxo gmhy cafw qbqo",
        "server": "smtp.gmail.com", 
        "port": 587
    }
}

PAGE_TITLE = "Apex Payroll System"
COMPANY_NAME = "Apex seekers edtech private limited"
COMPANY_ADDRESS = "Guna Complex, NewNo.443 & 445, Old No 304 & 305,\n1st Floor, Anna Salai, Teynampet, Chennai - 600 018."
LOGO_FILENAME = "Apex Seekers Logo.png" 

MARGIN = 25
BOX_HEIGHT = 480 
CONTENT_WIDTH = A4[0] - (2 * MARGIN)
app = FastAPI(
    title="Apex Payroll API",
    description="API for generating and emailing payslips from an uploaded payroll file.",
    version="1.0.0"
)

def format_currency(amount):
    """Format numbers to integer strings (no decimals) or 0."""
    try:
        if pd.isna(amount) or amount == "" or str(amount).strip() == "-":
            return "0"
        val = float(amount)
        return f"{int(round(val))}" 
    except (ValueError, TypeError):
        return "0"

def safe_get(row, col_name, default="-"):
    val = row.get(col_name, default)
    if pd.isna(val) if isinstance(val, (int, float, str)) else val is None or val == "":
        return default
    return str(val)

def draw_static_elements(canvas, doc):
    """
    Draws the fixed Black Border and the Footer Disclaimer.
    """
    canvas.saveState()
    x = MARGIN
    y = A4[1] - MARGIN - BOX_HEIGHT
    w = CONTENT_WIDTH
    h = BOX_HEIGHT
    
    canvas.setStrokeColor(colors.black)
    canvas.setLineWidth(2.5) 
    canvas.rect(x, y, w, h)
    canvas.setFont("Helvetica", 9)
    canvas.drawCentredString(A4[0] / 2, y - 15, 
                            "This is a computer-generated pay slip and does not require a signature or any company seal.")
    
    canvas.restoreState()

def generate_payslip_pdf(data):
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(buffer, pagesize=A4, 
                            leftMargin=MARGIN + 5, rightMargin=MARGIN + 5,
                            topMargin=MARGIN + 5, bottomMargin=MARGIN)
    
    elements = []
    styles = getSampleStyleSheet()
    
    style_comp_name = ParagraphStyle('CN', parent=styles['Heading1'], alignment=TA_CENTER, fontSize=14, fontName='Helvetica-Bold', spaceAfter=2)
    style_comp_addr = ParagraphStyle('CA', parent=styles['Normal'], alignment=TA_CENTER, fontSize=10, leading=12)
    style_month = ParagraphStyle('M', parent=styles['Normal'], alignment=TA_CENTER, fontSize=10, textTransform='uppercase', spaceBefore=4)
    
    if os.path.exists(LOGO_FILENAME):
        logo_obj = Image(LOGO_FILENAME, width=1.3*inch, height=0.7*inch)
        logo_obj.hAlign = 'LEFT'
    else:
        logo_obj = Paragraph("<b>LOGO<br/>MISSING</b>", styles['Normal'])

    header_text = [
        Paragraph(COMPANY_NAME, style_comp_name),
        Paragraph(COMPANY_ADDRESS, style_comp_addr),
        Paragraph(f"(PAYSLIP FOR THE MONTH OF {datetime.datetime.now().strftime('%B %Y').upper()} )", style_month)
    ]
    
    t_header = Table([[logo_obj, header_text]], colWidths=[1.5*inch, 5.5*inch])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'CENTER'),
        ('LEFTPADDING', (0,0), (0,0), 0),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 0.15*inch))
    
    doj = safe_get(data, 'DOJ')
    emp_data = [
        ["EMP id", safe_get(data, 'EMP_ID'), "Name", safe_get(data, 'NAME')],
        ["Designation", safe_get(data, 'DESIGNATION'), "Department", safe_get(data, 'DEPARTMENT')],
        ["Date of joining", doj, "Location", safe_get(data, 'LOCATION')],
        ["UAN no", safe_get(data, 'UAN'), "PAN no", safe_get(data, 'PAN')],
        ["ESIC no", safe_get(data, 'ESIC'), "Bank a/c no", safe_get(data, 'BANK_AC_NO')],
        ["Paid days", safe_get(data, 'PAID_DAYS'), "Lop days", safe_get(data, 'LOP_DAYS')],
        ["Leave taken", safe_get(data, 'LEAVE_TAKEN'), "Bal leave", safe_get(data, 'BAL_LEAVE')],
    ]
    
    t_emp = Table(emp_data, colWidths=[1.3*inch, 2.5*inch, 1.2*inch, 2.3*inch])
    t_emp.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2), 
        ('TOPPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(t_emp)
    elements.append(Spacer(1, 0.1*inch))
    
    earning_rows = [
        ("BASIC", 'BASIC_FIXED', 'BASIC_EARNED'),
        ("HRA", 'HRA_FIXED', 'HRA_EARNED'),
        ("CONVEYANCE ALLOWANCE", 'CONVEYANCE_FIXED', 'CONVEYANCE_EARNED'),
        ("MEDICAL REIMBURSE", 'MEDICAL_FIXED', 'MEDICAL_EARNED'),
        ("LEAVE TRAVEL ALLOWANCE", 'LTA_FIXED', 'LTA_EARNED'),
        ("SPECIAL ALLOWANCE", 'SPECIAL_FIXED', 'SPECIAL_EARNED'),
        ("INCENTIVE", None, 'INCENTIVE'),
        ("CLAIM", None, 'CLAIM'),
        ("ON DUTY", None, 'ON_DUTY'),
        ("OTHER EARNINGS", None, 'OTHER_EARNINGS')
    ]
    
    deduction_rows = [
        ("PF AMOUNT", 'PF_AMOUNT'),
        ("ESIC", 'ESIC_DED'),
        ("PROFESSIONAL TAX", 'PROF_TAX'),
        ("OTHER DEDUCTION", 'OTHER_DED')
    ]
    
    fin_data = [["EARNINGS", "Fixed", "Earned", "DEDUCTIONS", "Amount"]]
    
    max_len = max(len(earning_rows), len(deduction_rows))
    
    for i in range(max_len):
        row = []
        if i < len(earning_rows):
            lbl, k_fix, k_earn = earning_rows[i]
            v_fix = format_currency(data.get(k_fix)) if k_fix else ""
            v_earn = format_currency(data.get(k_earn))
            row.extend([lbl, v_fix, v_earn])
        else:
            row.extend(["", "", ""])
            
        if i < len(deduction_rows):
            lbl, k_amt = deduction_rows[i]
            v_amt = format_currency(data.get(k_amt))
            row.extend([lbl, v_amt])
        else:
            row.extend(["", ""])
        fin_data.append(row)
    
    fin_data.append([
        "GROSS TOTAL", "", format_currency(data.get('GROSS_TOTAL')), 
        "DEDUCTION TOTAL", format_currency(data.get('DEDUCTION_TOTAL'))
    ])
    
    col_w = [2.4*inch, 0.9*inch, 0.9*inch, 2.4*inch, 0.9*inch]
    
    t_fin = Table(fin_data, colWidths=col_w)
    t_fin.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), 
        ('FONTSIZE', (0,0), (-1,-1), 10),
        
        ('LINEABOVE', (0,0), (-1,0), 1.5, colors.black), 
        ('LINEBELOW', (0,0), (-1,0), 1.5, colors.black),
        
        ('ALIGN', (1,0), (2,-1), 'RIGHT'), 
        ('ALIGN', (4,0), (4,-1), 'RIGHT'), 
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('LINEABOVE', (0,-1), (-1,-1), 1.5, colors.black), 
        ('LINEBELOW', (0,-1), (-1,-1), 1.5, colors.black),
        
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_fin)
    
    elements.append(Spacer(1, 0.2*inch))
    
    net_val = format_currency(data.get('NET_PAY'))
    words = safe_get(data, 'NET_PAY_IN_WORDS', '')
    
    net_data = [
        ["NET PAY", net_val],
        ["In Words", f":{words}"]
    ]
    
    t_net = Table(net_data, colWidths=[1.5*inch, 6.0*inch])
    t_net.setStyle(TableStyle([
        ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'), 
        ('FONTNAME', (0,1), (0,1), 'Helvetica-Bold'), 
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ALIGN', (1,0), (1,0), 'LEFT'), 
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_net)
    
    doc.build(elements, onFirstPage=draw_static_elements, onLaterPages=draw_static_elements)
    buffer.seek(0)
    return buffer

def send_email_with_attachment(to_email, employee_name, pdf_buffer):
    smtp_conf = MOCK_SECRETS["smtp"]
    if 'email' not in smtp_conf or 'password' not in smtp_conf:
        return False, "SMTP configuration (email/password) is missing."

    msg = MIMEMultipart()
    msg['From'] = smtp_conf['email']
    msg['To'] = to_email
    msg['Subject'] = f"Payslip - {datetime.datetime.now().strftime('%B %Y')}"
    body = f"Dear {employee_name},\n\nPlease find attached your payslip for {datetime.datetime.now().strftime('%B %Y')}.\n\nBest Regards,\n{COMPANY_NAME}"
    msg.attach(MIMEText(body, 'plain'))

    pdf_attachment = MIMEApplication(pdf_buffer.getvalue(), _subtype="pdf")
    safe_name = "".join([c for c in employee_name if c.isalpha() or c.isdigit() or c==' ']).strip().replace(" ", "_")
    pdf_attachment.add_header('Content-Disposition', 'attachment', filename=f"Payslip_{safe_name}.pdf")
    msg.attach(pdf_attachment)

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(smtp_conf['server'], smtp_conf['port']) as server:
            server.starttls(context=context)
            server.login(smtp_conf['email'], smtp_conf['password'])
            server.send_message(msg)
        return True, "Success"
    except Exception as e:
        return False, str(e)

@app.post("/generate-and-email-payslips/")
async def generate_and_email(file: UploadFile = File(...)):
    """
    Uploads a payroll file (CSV or XLSX), generates payslips, and emails them 
    to the employees listed in the file.
    """
    if "smtp" not in MOCK_SECRETS or not all(k in MOCK_SECRETS["smtp"] for k in ["email", "password", "server", "port"]):
        raise HTTPException(
            status_code=500,
            detail="Server configuration error: SMTP secrets are missing or incomplete. Check MOCK_SECRETS in app.py."
        )
    try:
        file_extension = file.filename.split('.')[-1].lower()
        content = await file.read()
        file_buffer = io.BytesIO(content)

        if file_extension == 'csv':
            df = pd.read_csv(file_buffer)
        elif file_extension in ['xlsx', 'xls']:
            df = pd.read_excel(file_buffer)
        else:
            raise HTTPException(status_code=400, detail="Invalid file type. Only CSV and XLSX are supported.")
        
        df.columns = df.columns.str.strip()

    except Exception as e:
        raise HTTPException(status_code=400, detail=f"File processing error: {e}")

    required = ["NAME", "EMAIL", "NET_PAY", "EMP_ID"]
    if any(c not in df.columns for c in required):
        missing = [c for c in required if c not in df.columns]
        raise HTTPException(status_code=400, detail=f"Missing required columns in the file: {', '.join(missing)}")
    
    if 'DOJ' in df.columns:
        df['DOJ'] = pd.to_datetime(df['DOJ'], errors='coerce').dt.strftime('%m/%d/%Y').fillna('')

    s, f = 0, 0
    logs = []

    payroll_data = df.to_dict('records')

    for i, row in enumerate(payroll_data):
        name = str(row.get('NAME', 'N/A'))
        email = str(row.get('EMAIL', ''))
        
        if "@" not in email:
            logs.append({"status": "Skipped", "employee": name, "reason": "Bad Email"})
            f += 1
            continue
        
        try:
            pdf_buffer = generate_payslip_pdf(row)
            ok, msg = send_email_with_attachment(email, name, pdf_buffer)
            if ok:
                logs.append({"status": "Sent", "employee": name, "email": email})
                s += 1
            else:
                logs.append({"status": "Failed", "employee": name, "reason": msg})
                f += 1
        except Exception as e:
            logs.append({"status": "Error", "employee": name, "reason": str(e)})
            f += 1

    return JSONResponse(content={
        "status": "Complete",
        "total_records": len(df),
        "sent_count": s,
        "failed_count": f,
        "logs": logs
    })

if __name__ == "__main__":
    if not os.path.exists(LOGO_FILENAME):
        print(f"!!! WARNING: {LOGO_FILENAME} is missing. PDF will show 'LOGO MISSING'.")

    uvicorn.run(app, host="0.0.0.0", port=8000)