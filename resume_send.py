import smtplib
import pandas as pd
from email.message import EmailMessage
import os
import time

# Gmail credentials
GMAIL_USER = 'humber.resume.test@gmail.com'
GMAIL_APP_PASSWORD = 'flnzjjgrgrzgfdca'  # Use App Passwords for Gmail

# Resume versions with corresponding aliases

RESUME_VERSIONS = {
    'expedu': ('imadul.expedu@gmail.com', "/Users/imadulislamchowdhury/Downloads/humber_resume_test/resume_exp_edu.pdf")
}
'''
RESUME_VERSIONS = {
    'expedu': ('imadul.expedu@gmail.com', "/Users/imadulislamchowdhury/Downloads/humber_resume_test/resume_exp_edu.pdf"),
    'expnoedu': ('imadul.expnoedu@gmail.com', "/Users/imadulislamchowdhury/Downloads/humber_resume_test/resume_exp_noedu.pdf"),
    'noexpedu': ('imadul.noexpedu@gmail.com', "/Users/imadulislamchowdhury/Downloads/humber_resume_test/resume_noexp_edu.pdf"),
    'noexpnoedu': ('imadul.noexpnoedu@gmail.com', "/Users/imadulislamchowdhury/Downloads/humber_resume_test/resume_noexp_noedu.pdf")
}
'''

# Load recipients from Excel
df = pd.read_excel('/Users/imadulislamchowdhury/Downloads/humber_resume_test/recipients.xlsx')  # Columns: Email, Company, Position


def send_email(to_email, from_alias, resume_path, company, position):
    msg = EmailMessage()

    msg['Subject'] = f"Application for {position} at {company}"
    msg['From'] = f"Imadul Chowdhury <{from_alias}>"
    msg['To'] = to_email

    msg.set_content(f"""
Dear Hiring Manager at {company},

I am writing to express my interest in the {position} role at your organization.

Please find my resume attached for your consideration.

Best regards,  
Imadul Islam Chowdhury
""")

    # Attach resume
    with open(resume_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(resume_path)
        msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    # Send email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        smtp.send_message(msg)
        print(f"Sent from {from_alias} to {to_email} for {company} - {position}")


# Loop over each recipient and send all 4 versions
for index, row in df.iterrows():
    to_email = row['Email']
    company = row['Company']
    position = row['Position']

    for tag, (from_alias, resume_path) in RESUME_VERSIONS.items():
        send_email(to_email, from_alias, resume_path, company, position)
        time.sleep(30)  # 30-second pause to avoid spam flagging
