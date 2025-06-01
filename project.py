import imaplib
import email
from email.header import decode_header
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

class EmailParser:
    def __init__(self, root):
        """
        Initializes the main application window and builds the GUI.
        """
        self.root = root
        self.root.title("Email Fetcher")

        self.build_gui()  # Call method to construct GUI components

    def build_gui(self):
        """
        Sets up all GUI components (labels, entries, buttons, treeview).
        """
        # Label and entry for Email input
        tk.Label(self.root, text="Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.email_entry = tk.Entry(self.root, width=40)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5)

        # Label and entry for Password input (masked with "*")
        tk.Label(self.root, text="Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.password_entry = tk.Entry(self.root, show="*", width=40)
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        # Button to trigger fetching emails when clicked
        fetch_button = tk.Button(self.root, text="Fetch Last 20 Emails", command=self.fetch_emails)
        fetch_button.grid(row=2, column=0, columnspan=2, pady=10)

        # Treeview widget to display fetched emails in tabular form
        columns = ("From", "Subject", "Category")  # Table columns
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=10)

        # Set the headings and column widths for the treeview
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.W, width=250)

        self.tree.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

    def fetch_emails(self):
        """
        Handles the entire process of connecting to the email server,
        logging in, fetching the last 20 emails, and displaying them.
        """
        # Get the email and password entered by the user in the GUI
        email_address = self.email_entry.get()
        password = self.password_entry.get()

        # Validate input fields - both must be filled
        if not email_address or not password:
            messagebox.showwarning("Input Error", "Please enter both email and password")
            return

        IMAP_SERVER = "imap.gmail.com"  # Gmail's IMAP server

        try:
            # Create a secure SSL connection to the Gmail IMAP server
            mail = imaplib.IMAP4_SSL(IMAP_SERVER)
            # Attempt to login with provided credentials
            mail.login(email_address, password)
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
        # "ALL" means to search all emails in the mailbox.
        # status: a string like "OK" or "NO" indicating if the search succeeded.
        # messages: a list (usually with one string) containing space-separated email IDs that match the search.

        if status != "OK":  # This line checks if the search failed.
            messagebox.showerror("Error", "Failed to retrieve emails")
            return

        email_ids = messages[0].split()
        # messages[0] is a byte string of email IDs separated by spaces.
        # .split() splits the string by spaces into a list of individual email IDs.

        email_data = []  # List to store parsed email info

        # Loop through the last 20 email IDs (most recent emails)
        for eid in email_ids[-20:]:
            # Fetch the full raw email message by ID
            res, msg = mail.fetch(eid, "(RFC822)")
            # res stores the response status (e.g., "OK" or "NO").
            # msg is a list containing the raw email data in bytes.
            # "(RFC822)" means fetch the entire raw email message.

            for response in msg:
                # Check if the current response is a tuple (it contains the raw email data)
                if isinstance(response, tuple):
                    # Parse the raw bytes to an email message object
                    msg_obj = email.message_from_bytes(response[1])

                    # Decode the email subject header which may be encoded
                    subject, encoding = decode_header(msg_obj["Subject"])[0]
                    # decode_header() is used because the subject can be encoded (e.g., UTF-8, ISO-8859-1)
                    if isinstance(subject, bytes):
                        # Decode bytes to string using detected encoding or fallback utf-8
                        subject = subject.decode(encoding or "utf-8")

                    # Get the sender's email from the "From" header
                    from_ = msg_obj.get("From")

                    # Categorize emails based on keywords in the subject
                    category = self.categorize_email(subject)

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

        # Display the email data in the treeview
        self.display_emails(email_data)

    def categorize_email(self, subject):
        """
        Determines the category of an email based on keywords in the subject line.
        """
        subject_lower = subject.lower()
        if "invoice" in subject_lower:
            return "Finance"
        elif "meeting" in subject_lower:
            return "Work"
        elif "promo" in subject_lower:
            return "Promotions"
        else:
            return "Others"

    def display_emails(self, email_data):
        """
        Clears the treeview and inserts new email data rows.
        """
        # Clear any existing rows in the treeview before inserting new data
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Insert each email's info into the treeview table
        for item in email_data:
            self.tree.insert("", tk.END, values=(item["From"], item["Subject"], item["Category"]))

# Entry point of the application
if __name__ == "__main__":
    root = tk.Tk()              # Create the main Tkinter window
    app = EmailParser(root) # Instantiate the app with the root window
    root.mainloop()             # Start the Tkinter event loop
