import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import make_msgid

# ========== CONFIGURATION ==========
EXCEL_FILE = "students.xlsx"
TEMPLATE = "certificate_template.png"
OUTPUT_DIR = "certificates"
GMAIL_ADDRESS = "poojithagavara4@gmail.com"  # 🔁 Replace with your Gmail address
APP_PASSWORD = "azcy avks rnfw tmfp" \
""  # 🔁 Replace with App Password (no spaces)

# ========== SETUP ==========
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# Load Excel file
try:
    students = pd.read_excel(EXCEL_FILE)
except Exception as e:
    print(f"❌ Failed to load Excel file: {e}")
    exit()

# Validate required columns
required_columns = {"Name", "Roll number", "Department", "Email"}
if not required_columns.issubset(students.columns):
    print("❌ Excel file must contain columns: Name, Roll number, Department, Email")
    exit()

# ========== CERTIFICATE GENERATOR ==========
def generate_certificate(name, roll, dept):
    file_path = os.path.join(OUTPUT_DIR, f"{roll}_{name}.pdf")
    c = canvas.Canvas(file_path, pagesize=A4)

    if os.path.exists(TEMPLATE):
        c.drawImage(TEMPLATE, 0, 0, width=A4[0], height=A4[1])
    else:
        print(f"⚠ Certificate template '{TEMPLATE}' not found. Skipping background.")

    c.setFont("Helvetica-Bold", 28)
    c.drawCentredString(A4[0]/2, A4[1]/2 + 50, name)

    c.setFont("Helvetica", 16)
    c.drawCentredString(A4[0]/2, A4[1]/2, f"Roll No: {roll} | Dept: {dept}")

    c.save()
    return file_path

# ========== SEND EMAIL ==========
def send_email(to_email, subject, body, attachment_path):
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = GMAIL_ADDRESS
        msg['To'] = to_email
        msg.set_content("This is a MIME-aware email. If you see this, your client doesn't support HTML.")
        msg.add_alternative(body, subtype='html')

        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)

        msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(GMAIL_ADDRESS, APP_PASSWORD)
            smtp.send_message(msg)

        print(f"✅ Email sent to {to_email}")

    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")

# ========== PROCESS STUDENTS ==========
for _, row in students.iterrows():
    try:
        name = row["Name"]
        roll = row["Roll number"]
        dept = row["Department"]
        email = row["Email"]

        pdf_path = generate_certificate(name, roll, dept)

        subject = "Certificate of Participation"
        body = f"""
        <html>
        <body>
        <p>Dear {name},</p>
        <p>Please find attached your certificate of participation.</p>
        <p>Best regards,<br>Team</p>
        </body>
        </html>
        """

        send_email(email, subject, body, pdf_path)

    except Exception as e:
        print(f"❌ Error processing student {row.get('Name', 'Unknown')}: {e}")
