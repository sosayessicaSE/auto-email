import os
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

SMTP_SERVER = os.getenv('SMTP_SERVER')  
SMTP_PORT = int(os.getenv('SMTP_PORT', 587)) 
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')

df = pd.read_excel('books.xlsx', sheet_name='bingo')
mails = df['mail'].tolist()
names = df['name'].tolist()


BINGO_FOLDER = 'bingo_images'
bingo_images = os.listdir(BINGO_FOLDER)
random.shuffle(bingo_images)

def create_server_connection():
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    print("Login successful!")
    return server

def send_email(server, to_email, subject, body, image_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with open(image_path, 'rb') as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header('Content-Disposition', 'attachment', filename=os.path.basename(image_path))
            msg.attach(mime_img)
        
        server.send_message(msg)
        print(f"Email sent to {to_email} with card {os.path.basename(image_path)}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

subject = "Your Bingo Card 🎯"
message_template = """
Hi {name},

Attached is your unique bingo card. Enjoy!
"""

server = create_server_connection()

for name, mail in zip(names, mails):
    if mail and bingo_images:
        image_file = bingo_images.pop()  
        image_path = os.path.join(BINGO_FOLDER, image_file)
        
        personalized_message = message_template.format(name=name)
        send_email(server, mail, subject, personalized_message, image_path)
    else:
        print(f"Invalid email for {name} or no more images to send!!")

server.quit()
