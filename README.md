📄 FastAPI Payslip Generator & Email Sender
     A powerful and automated backend system built using FastAPI to generate custom-designed PDF payslips and email them directly to employees. This API processes payroll data, formats it into a professional A4 payslip using ReportLab, and sends personalized PDFs through an SMTP mail server.

✨ Features
⚡ High-Performance FastAPI Backend
     Asynchronous file processing
     Automatic PDF generation for each employee
     Email delivery via SMTP

📁 Upload & Process Payroll Files
     Accepts .xlsx or .csv formats
     Extracts employee and salary details row-by-row

🧾 Custom PDF Payslip Template
     Designed using ReportLab
     A4 layout with thick black border
     Includes company header, employee details, and salary breakdown

📧 Automated Email Sending
     SMTP integration (Gmail, Outlook, Company Servers, etc.)
     Sends one personalized payslip per employee

📧 SMTP Configuration (Important)
     Inside app.py, locate the following dictionary:
     MOCK_SECRETS = {
        "smtp": {
           "email": "YOUR_SENDER_EMAIL@example.com",
           "password": "YOUR_APP_PASSWORD",
           "server": "smtp.gmail.com",
           "port": 587
       }
   }

🚀 API Usage
🔹 Endpoint
      POST /generate-and-email-payslips/
🔹 Input
      Multipart Form Data containing:
         An uploaded .xlsx or .csv payroll file.
🔹 Output
      JSON summarizing:
         Total records processed
         Successful emails
         Failed emails
         Per-employee log status
