import imaplib
import email
import re
from textblob import TextBlob
import pandas as pd
import os
import bs4  # For HTML parsing

# Gmail IMAP credentials
EMAIL_ACCOUNT = 'imadul.expedu@gmail.com'
EMAIL_PASSWORD = 'mbym flyj qowf xxuq'  # Use your App Password

# Output file
OUTPUT_FILE = '/Users/imadulislamchowdhury/Downloads/humber_resume_test/responses.xlsx'

# Load existing data or create new DataFrame
if os.path.exists(OUTPUT_FILE):
    df = pd.read_excel(OUTPUT_FILE)
else:
    df = pd.DataFrame(columns=['From Email', 'Company', 'Position', 'Response'])

# Connect to Gmail IMAP
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
mail.select('inbox')

# Search for all UNSEEN emails
result, data = mail.search(None, 'UNSEEN')

for num in data[0].split():
    result, msg_data = mail.fetch(num, '(RFC822)')
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    # Extract From Email
    from_email = email.utils.parseaddr(msg['From'])[1]

    # Extract Subject and parse Company & Position
    subject = msg['Subject']
    company = "Unknown"
    position = "Unknown"

    # Use regex for robust parsing
    # Expected subject format: Re: Application for Data Analyst at TechCorp
    clean_subject = subject.replace("Re:", "").strip()
    match = re.search(r'Application for (.*?) at (.*)', clean_subject)

    if match:
        position = match.group(1).strip()
        company = match.group(2).strip()

    # Extract body with fallback to HTML
    body = None

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                body = part.get_payload(decode=True).decode(errors='ignore').strip()
                break
        if not body:
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    html_content = part.get_payload(decode=True).decode(errors='ignore')
                    soup = bs4.BeautifulSoup(html_content, "html.parser")
                    body = soup.get_text().strip()
                    break
    else:
        body = msg.get_payload(decode=True).decode(errors='ignore').strip()

    # Handle empty body
    if not body:
        body = "(No readable content found)"

    #Sentiment Analysis
    # Lowercase body for intent matching
    body_lower = body.lower()

    # Intent-based classification
    positive_keywords = [
        "interview",
        "next step",
        "proceed",
        "availability",
        "schedule a call",
        "move forward with you",
        "invited to interview",
        "call to discuss",
        "set up a meeting",
        "shortlisted",
        "selected for the next round",
        "advance your application",
        "look forward to speaking",
        "we would like to meet",
        "pre-screening call",
        "initial call",
        "let's connect",
        "looking forward to talking",
        "considering you for the role",
        "moving you forward",
        "let's proceed",
        "move ahead with your candidacy"
    ]

    negative_keywords = [
        "unfortunately",
        "another applicant",
        "we have decided to move forward with",
        "not moving forward",
        "reject",
        "decline",
        "position has been filled",
        "better suited candidates",
        "regret to inform",
        "we are unable to proceed",
        "not selected",
        "after careful consideration",
        "we will not be progressing",
        "not pursuing your candidacy",
        "we are closing this application",
        "position is no longer available",
        "considered other candidates",
        "decided not to proceed",
        "wish you all the best in your search",
        "no longer under consideration"
    ]

    response = "Neutral/Unknown"  # Default fallback

    if any(kw in body_lower for kw in positive_keywords):
        response = "Positive"
    elif any(kw in body_lower for kw in negative_keywords):
        response = "Negative"
    else:
        # Fallback to TextBlob sentiment
        blob = TextBlob(body)
        polarity = blob.sentiment.polarity

        if polarity > 0.1:
            response = "Positive (by sentiment)"
        elif polarity < -0.1:
            response = "Negative (by sentiment)"

    # Append result to DataFrame
    df = pd.concat([df, pd.DataFrame([{
        'From Email': from_email,
        'Company': company,
        'Position': position,
        'Response': response
    }])], ignore_index=True)

    print(f"Processed reply from {from_email}: {company} - {position} => {response}")

# Save to Excel
df.to_excel(OUTPUT_FILE, index=False)
print(f"Results saved to {OUTPUT_FILE}")
