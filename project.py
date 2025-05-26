import imaplib
import email
from email.header import decode_header
import pandas as pd

# Email account credentials
EMAIL = "your_email@example.com"
PASSWORD = "your_password"
IMAP_SERVER = "imap.gmail.com"

# Connect to the server and login
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, PASSWORD)

# Select the inbox
mail.select("inbox")

# Search for all emails
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()

# Store parsed emails
email_data = []

for eid in email_ids[-10:]:  # fetch last 10 emails for example
    res, msg = mail.fetch(eid, "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])

            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding or "utf-8")
            
            from_ = msg.get("From")

            # Categorization based on keywords
            if "invoice" in subject.lower():
                category = "Finance"
            elif "meeting" in subject.lower():
                category = "Work"
            elif "promo" in subject.lower():
                category = "Promotions"
            else:
                category = "Others"

            email_data.append({
                "From": from_,
                "Subject": subject,
                "Category": category
            })

# Convert to DataFrame
df = pd.DataFrame(email_data)
print(df)
