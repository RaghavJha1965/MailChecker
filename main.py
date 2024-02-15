import imaplib
import email
from email.header import decode_header
import webbrowser
import os

# Account credentials
username = "testRV2005@outlook.com"
password = "R@ghav132"  # Use the App Password you generated
imap_server = "outlook.office365.com"

def clean(text):
    # Clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

# Number of top emails to fetch
N = 3

# Create an IMAP4 class with SSL
imap = imaplib.IMAP4_SSL(imap_server)

# Authenticate using the App Password
imap.login(username, password)

# Select the inbox mailbox
status, messages = imap.select("INBOX")

# Total number of emails
messages = int(messages[0])

# Adjust the starting message ID
start_message_id = max(messages - N + 1, 1)

# Iterate over the messages
for i in range(messages, start_message_id - 1, -1):
    try:
        # Fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")

        # Process the email message
        for response in msg:
            if isinstance(response, tuple):
                # Parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # Decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)
                # Decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                print("Subject:", subject)
                print("From:", From)
                # Rest of your code to process the email message...
            print("=" * 100)
    except imaplib.IMAP4.error as e:
        print(f"Error fetching message {i}: {e}")

# Close the connection and logout
imap.close()
imap.logout()
