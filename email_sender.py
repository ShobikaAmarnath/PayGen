# email_sender.py
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email_config import SMTP_SERVERS

def connect_to_server(sender_email, sender_password):
    """
    Detects the email provider, connects to the SMTP server, and logs in.
    Returns the active server connection object.
    """
    # Auto-detect the email provider
    provider = 'microsoft' # Default to Microsoft for custom domains
    if 'gmail.com' in sender_email:
        provider = 'gmail'
    
    server_config = SMTP_SERVERS[provider]
    smtp_server, smtp_port = server_config['server'], server_config['port']
    print(f"Detected provider: {provider.title()}. Connecting to {smtp_server}...")

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        print("Connection and login successful.")
        return server
    except Exception as e:
        print(f"Failed to connect or log in: {e}")
        return None

def send_single_email(server, sender_email, recipient_email, employee_name, period, pdf_path):
    """Creates and sends a single email using an existing server connection."""
    try:
        # --- Email Content ---
        subject = f"Your Payslip for {period.strftime('%B %Y')}"
        html_body = f"""
        <html>
          <body>
            <p>Dear {employee_name},</p>
            <p>Please find your payslip for <b>{period.strftime('%B %Y')}</b> attached to this email.</p>
            <p>If you have any questions, please contact the HR department.</p>
            <br>
            <p>Best Regards,</p>
            <p><b>HR Department</b></p>
          </body>
        </html>
        """
        
        # --- Create the Email Message ---
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(html_body, 'html'))

        # --- Attach the PDF File ---
        with open(pdf_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {pdf_path.name}")
        msg.attach(part)
        
        # --- Send using the existing connection ---
        server.send_message(msg)
        print(f"Successfully sent email to {recipient_email}")
        return True, "Email sent successfully!"
    except Exception as e:
        error_msg = f"Failed to send email to {recipient_email}: {e}"
        print(error_msg)
        return False, error_msg

def close_connection(server):
    """Closes the SMTP server connection."""
    if server:
        server.quit()
        print("Connection closed.")