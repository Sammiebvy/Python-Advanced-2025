import imaplib
import email
from email.header import decode_header
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

def fetch_emails():
    # Get the email and password entered by the user in the GUI
    EMAIL = email_entry.get()
    PASSWORD = password_entry.get()

    # Validate input fields - both must be filled
    if not EMAIL or not PASSWORD:
        messagebox.showwarning("Input Error", "Please enter both email and password")
        return

    IMAP_SERVER = "imap.gmail.com"  # Gmail's IMAP server

    try:
        # Create a secure SSL connection to the Gmail IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        # Attempt to login with provided credentials
        mail.login(EMAIL, PASSWORD)
    # Handle authentication failures like wrong email or password
    except imaplib.IMAP4.error as e:
        messagebox.showerror("Login Failed", f"Invalid email or password.\nDetails: {e}")
        return
    # Catch any other unexpected errors during connection or login
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred.\nDetails: {e}")
        return

    # Select the "inbox" mailbox to fetch emails from
    mail.select("inbox")

    # Search for all email messages in the inbox
    status, messages = mail.search(None, "ALL") 
    #mail is an IMAP connection object connected to an email server. Search is a method in the IMAP library used to find an email in a certain criteria
    #"All" means to search all emails in the mailbox
    # status: a string like "OK" or "NO" indicating if the search succeeded.
    # messages: a list (usually with one string) containing space-separated email IDs that match the search.
    if status != "OK": #This line checks if the search failed.
        messagebox.showerror("Error", "Failed to retrieve emails")
        return

    # messages[0] is a byte string of email IDs separated by spaces
    email_ids = messages[0].split()
    # messages is a list with one string element containing all matching email IDs separated by spaces.
    # messages[0] gets that single string.
    # .split() splits the string by spaces into a list of individual email IDs.
    # So email_ids becomes a list of email IDs (each a string), e.g., ['1', '2', '3', ...].

    email_data = []  # List to store parsed email info

    # Loop through the last 10 email IDs (most recent emails)
    for eid in email_ids[-10:]:
        # Fetch the full raw email message by ID
        res, msg = mail.fetch(eid, "(RFC822)")
        #using the fetch method on the mail object to retrieve the full email by its ID (eid).
        # res stores the response status (e.g., "OK" or "NO").
        # msg is a list containing the raw email data in bytes.
        # "(RFC822)" means fetch the entire raw email message.
        for response in msg:
            # Check if the current response is a tuple (the tuple contains the raw email data).
            # This is necessary because the msg list can sometimes contain other data types (e.g., strings).
            if isinstance(response, tuple):
                # Parse the raw bytes to an email message object
                msg_obj = email.message_from_bytes(response[1])

                # Decode the email subject header which may be encoded
                subject, encoding = decode_header(msg_obj["Subject"])[0]
                # Extract the "Subject" header from the email.
                # decode_header() is used because the subject can be encoded (e.g., UTF-8, ISO-8859-1).
                # It returns a list of tuples (decoded_string, encoding), so [0] gets the first part.
                if isinstance(subject, bytes):
                    # Decode bytes to string using detected encoding or fallback utf-8
                    subject = subject.decode(encoding or "utf-8")

                # Get the sender's email from the "From" header
                from_ = msg_obj.get("From")

                # Categorize emails based on keywords in the subject
                if "invoice" in subject.lower():
                    category = "Finance"
                elif "meeting" in subject.lower():
                    category = "Work"
                elif "promo" in subject.lower():
                    category = "Promotions"
                else:
                    category = "Others"

                # Append the extracted info as a dictionary to the list
                email_data.append({
                    "From": from_,
                    "Subject": subject,
                    "Category": category
                })

    # If no emails were found/displayed, inform the user
    if not email_data:
        messagebox.showinfo("No Emails", "No emails found.")
        return

    # Clear any existing rows in the treeview before inserting new data
    for row in tree.get_children():
        tree.delete(row)

    # Insert each email's info into the treeview table
    for email_item in email_data:
        tree.insert("", tk.END, values=(email_item["From"], email_item["Subject"], email_item["Category"]))

# Set up the main Tkinter window
root = tk.Tk()
root.title("Email Fetcher")

# Label and entry for Email input
tk.Label(root, text="Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
email_entry = tk.Entry(root, width=40)
email_entry.grid(row=0, column=1, padx=5, pady=5)

# Label and entry for Password input (masked with "*")
tk.Label(root, text="Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
password_entry = tk.Entry(root, show="*", width=40)
password_entry.grid(row=1, column=1, padx=5, pady=5)

# Button to trigger fetching emails when clicked
fetch_button = tk.Button(root, text="Fetch Last 10 Emails", command=fetch_emails)
fetch_button.grid(row=2, column=0, columnspan=2, pady=10)

# Treeview widget to display fetched emails in tabular form
columns = ("From", "Subject", "Category")  # Table columns
tree = ttk.Treeview(root, columns=columns, show="headings", height=10)

# Set the headings and column widths for the treeview
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor=tk.W, width=250)

tree.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

# Start the Tkinter event loop (keep window open)
root.mainloop()
