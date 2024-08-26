import email.utils
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List


def get_from_env(key: str) -> str:
    value = os.getenv(key)
    assert isinstance(value, str), f"{key} not set in .env"
    return value


def send_mail(
    subject: str,
    body: str,
    attachment_folder_path: str,
) -> bool:
    server: str = get_from_env("SMTP_SERVER")
    sender: str = get_from_env("SMTP_SENDER")
    recipients: str = get_from_env("SMTP_RECIPIENTS")
    recipients_lst: List[str] = recipients.split(";")

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipients
    msg["Date"] = email.utils.formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "html", "utf-8"))

    for file_name in os.listdir(attachment_folder_path):
        file_path = os.path.join(attachment_folder_path, file_name)
        if os.path.isdir(file_path):
            continue
        with open(file_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=file_name)
        part.add_header("Content-Disposition", "attachment", filename=file_name)
        msg.attach(part)

    try:
        with smtplib.SMTP(server, 25) as smtp:
            response = smtp.sendmail(sender, recipients_lst, msg.as_string())
            if response:
                print("Failed to send email to the following recipients:")
                for recipient, error in response.items():
                    print(f"{recipient}: {error}")
                return False
            else:
                print("Email sent successfully.")
                return True
    except smtplib.SMTPException as e:
        print(f"Failed to send email: {e}")
        return False
