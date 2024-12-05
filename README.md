Bingo Card Email Sender
This Python script sends personalized bingo cards to a list of recipients via email. The bingo cards are attached as images and the recipients are personalized with their names in the message.

Requirements
Before running the script, ensure you have the following installed:

Python 3.x
Pandas
dotenv
smtplib
email.mime

Open your terminal or command prompt and install the required dependencies:

pip install pandas python-dotenv

You need to set the following environment variables in a .env file to store sensitive information:

SMTP_SERVER=your_smtp_server
SMTP_PORT=your_smtp_port (default 587)
EMAIL_ADDRESS=your_email_address
EMAIL_PASSWORD=your_email_password

books.xlsx: An Excel file containing the names and email addresses of the recipients. The sheet should be named bingo with columns name and mail.
bingo_images/: A folder containing bingo card images. These images will be randomly assigned to the recipients.

How It Works
The script reads the books.xlsx file and retrieves the list of names and emails.
It loads the bingo images from the bingo_images/ folder and shuffles them.
For each recipient, it sends an email with a personalized message and a bingo card image attached.
The email's subject is "Your Bingo Card 🎯" and includes the recipient's name in the body.

Imports: The script uses libraries like os, random, smtplib, email.mime, dotenv, and pandas.
load_dotenv(): Loads environment variables from the .env file.
Pandas: Reads the Excel file and extracts the names and emails of recipients.
SMTP Setup: Connects to the SMTP server and logs in using the credentials from the .env file.
Email Composition: Creates a multipart email message, attaches the bingo card image, and sends the email to each recipient.
Randomization: Bingo cards are randomly shuffled before being assigned to recipients.

How to Run
Create a .env file with your SMTP server credentials.
Place the books.xlsx file with the recipient details and ensure the bingo card images are in the bingo_images folder.

Run the script:
python bingo_email_sender.py

Example .env File: 
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
EMAIL_ADDRESS=your_email@gmail.com
EMAIL_PASSWORD=your_email_password


Ensure your email account allows less secure apps or use an app-specific password for Gmail.
Make sure there are enough bingo card images for each recipient.
