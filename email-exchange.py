from jproperties import Properties
import win32com.client
import pythoncom
import smtplib
from email.mime.text import MIMEText


# Load config
with open("config.properties", "r+b") as f:
    p = Properties()
    p.load(f, "utf-8")

    smtp_server = p.properties.get("smtp.server")
    smtp_port = p.properties.get("smtp.port")
    smtp_user = p.properties.get("smtp.user")
    smtp_password = p.properties.get("smtp.password")
    recipient = p.properties.get("recipient")


class OutlookHandler:
    def __init__(self):
        try:
            self.outlook = win32com.client.DispatchWithEvents(
                "Outlook.Application", OutlookEvents
            )
        except Exception as e:
            print(f"Error: {e}")


class OutlookEvents:
    def OnNewMailEx(self, received_items_ids):
        for id in received_items_ids.split(","):
            mail = win32com.client.Dispatch(
                "Outlook.Application"
            ).Session.GetItemFromID(id)
            subject = mail.Subject
            body = mail.Body
            # print(mail)
            print("Receive:", subject, body)
            send_notification(subject, body)


def send_notification(subject, body):
    mimet = MIMEText(body)
    mimet["Subject"] = subject
    mimet["From"] = smtp_user
    mimet["To"] = recipient
    msg = mimet.as_string()

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.ehlo()
    server.starttls()
    server.login(smtp_user, smtp_password)
    status = server.sendmail(smtp_user, recipient, msg)
    if status == {}:
        print("OK")
    else:
        print("FAILED")
    server.quit()


if __name__ == "__main__":
    outlook_handler = OutlookHandler()
    while True:
        pythoncom.PumpWaitingMessages()
    # print("test")
