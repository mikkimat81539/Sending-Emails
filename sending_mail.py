"""Send the emails from the Excel file.

STEPS:
Step 1: Import the modules needed

pandas: Helps you read and work with the Excel file.
smtplib: Helps send the emails.
email: Helps format the emails (It’s not made to send emails or connect to servers like SMTP).

Step 2: Read the Excel File
Use pandas to read the Excel file

Step 3: Set Up Email Settings
You need to tell Python where to send the emails from.

For this, you’ll need:
Your email (for example, Gmail).
Your email password (be careful with this!).
The server address to send emails (e.g., Gmail's SMTP server is smtp.gmail.com).

The basic idea is to connect to your email provider's server and log in using your email and password.

Step 4: Create an Email for Each Person
You will loop through the rows in your Excel file (one row for each person).

For each person, you will create an email:
To: Their email address.
Subject: The subject of the email (like "Hello").
Body: The personalized message for that person.

Step 5: Send the email

Step 6. Error Handling
Make sure to check for problems, like if the email doesn’t send, or if there's an issue with the email address.

7. close email connection

After all emails are sent, you can close the email connection.

"""

import pandas as pd, smtplib as smtp

df = pd.read_excel("email.xlsx")

email1 = df.iloc[0, 1] # [row, column]
email2 = df.iloc[1, 1]

message1 = df.iloc[0, 2]
message2 = df.iloc[1, 2]

from_email = "test2@gmail.com"
password = "#### #### #### ####"
server_address = "smtp.gmail.com"
server_port = 587

with smtp.SMTP(server_address, server_port) as server:
    server.starttls()

    # Log in to the SMTP server
    server.login(from_email, password)

    for i in range(2):
        to_email = df.iloc[i, 1]
        messages = df.iloc[i, 2]

        subject = "Task for Employees (Python Test)."

        # Create the email headers and body (plain text format)
        setup = f"From: {from_email}\r"
        setup += f"To: {to_email}\r"
        setup += f"Subject: {subject}\r"  # Include the subject header
        setup += "\r"  # Separate headers from the body
        setup += messages  # Add the body of the email

        server.sendmail(from_addr=from_email, to_addrs=to_email, msg=setup.encode("utf-8"))

    server.quit()
