import pandas as pd
import qrcode
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText  # <-- Import this
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

# Step 1: Read the Excel file
file_path = '/home/kaleab/Downloads/myemail.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path)

# Step 2: Generate QR codes and send emails
def generate_qr_code(data, filename):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    img.save(filename)

def send_email(to_email, subject, body, attachment_path):
    from_email = 'dar@ati.gov.et'  # Replace with your email
    from_password = 'Di@#ati24'  # Replace with your email password

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={attachment_path}')
        msg.attach(part)

    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(from_email, from_password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    organization = row['organization']
    fullname = row['fullname']
    position = row['position']
    email = row['email']
    participant = row['participant']

    # Create a string with all the details
    details = f"Organization: {organization}\nFull Name: {fullname}\nPosition: {position}\nEmail: {email}\nParticipant: {participant}"

    # Generate QR code
    qr_filename = f"{fullname}_qrcode.png"
    generate_qr_code(details, qr_filename)

    # Send email with QR code attached
    subject = f"Check-In code for Digital Agriculture Roadmap - DAR Launching event participants"
    body = f"Dear {fullname} \n\n\nThis is your check-in ID for the Digital Agriculture Roadmap (DAR) Launch Event. Please bring this email or the attached QR code on your mobile device, as it will be scanned at the entrance to the venue.\nThank you for your cooperation.\n\n Regards,"
    send_email(email, subject, body, qr_filename)

    print(f"QR code generated and email sent for {fullname}")

print("All QR codes generated and emails sent successfully.")
