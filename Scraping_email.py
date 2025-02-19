import win32com.client
import pandas as pd
from datetime import datetime

# Define the Outlook object
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Define the root folder (usually the user's account)
root_folder = outlook.Folders.Item("samuel@samuel.com")  # Replace with your account name

# Access the "Inbox" folder
inbox_folder = root_folder.Folders.Item("Inbox")

# Access the specific subfolder, in this case, "Test"
Test = inbox_folder.Folders.Item("Test")  # Use the correct subfolder name here

# Retrieve messages from the folder defined above
messages = Test.Items

# Filter messages with the specific subject and sort them by received date (from newest to oldest)
filtered_messages = [message for message in messages if message.Class == 43 and "testing" in message.Subject]
filtered_messages = sorted(filtered_messages, key=lambda msg: msg.ReceivedTime, reverse=True)

# List to store email data
emails_data = []

# Iterate through only the first 10 filtered emails
for i, message in enumerate(filtered_messages):
    if i >= 10:  # Stop after collecting 10 emails
        break
    
    subject = message.Subject
    body = message.Body
    received_time = message.ReceivedTime
    sender = message.SenderName

    # Convert the date to a format compatible with pandas
    if received_time is not None:
        received_time = datetime(received_time.year, received_time.month, received_time.day, received_time.hour, received_time.minute, received_time.second)

    # Add data to the list
    emails_data.append([subject, body, received_time, sender])

# Create a DataFrame with the 10 collected emails
df = pd.DataFrame(emails_data, columns=["Subject", "Body", "Received Date", "Sender"])

# Save the data to an Excel file
df.to_excel("emails_filtered.xlsx", index=False)

print("The data of the 10 most recent filtered emails has been successfully saved in emails_filtered.xlsx")