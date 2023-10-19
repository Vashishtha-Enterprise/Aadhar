import win32com.client as win32
from datetime import datetime
import pandas as pd

def read_emails_from_folder(folder_path):
    # Connect to Outlook
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # # Split the folder path into separate folder names
    # folder_names = folder_path.split("/")
    
    # # Get the root folder
    # folder = namespace.GetDefaultFolder(6)
    
    # # Traverse the folder hierarchy
    # for folder_name in folder_names:
    #     folder = folder.Folders[folder_name]
    
    # # Retrieve all emails in the specified folder
    # emails = folder.Items
    
    # # Create a list to store email data
    # email_data = []
    
     
    # Prompt user to select the folder
    folder = namespace.PickFolder()
    
    # Retrieve all emails in the selected folder
    emails = folder.Items
    
    # Create a list to store email data
    email_data = []
    
    # Iterate over each email
    for email in emails:
        # Extract relevant email information
        subject = email.Subject
        sender_name = email.Sender.Name
        sender_email = email.Sender.Address
        received_time = email.ReceivedTime
        
        # Convert received_time to timezone-unaware datetime object
        received_time = received_time.replace(tzinfo=None)  # Remove timezone information
       
        # # Convert received_time to timezone-aware datetime object
        # received_time = received_time.replace(tzinfo=None)  # Remove timezone information
        # received_time = datetime.combine(received_time, datetime.min.time())  # Convert to datetime object
        # received_time = received_time.astimezone()  # Convert to timezone-aware datetime
        
        # Append email data to the list
        email_data.append([subject, sender_name, sender_email, received_time])
    
    # Create a pandas DataFrame from the email data
    df = pd.DataFrame(email_data, columns=["Subject", "Sender Name", "Sender Email", "Received Time"])
    
    return df

def save_emails_to_excel(df, filename):
    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False)

# Example usage
folder_path = "Inbox"  # Replace with the desired folder path or name
emails_df = read_emails_from_folder(folder_path)
save_emails_to_excel(emails_df, "emails.xlsx")