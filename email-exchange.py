from jproperties import Properties
import win32com.client
import pythoncom
import smtplib
from email.mime.text import MIMEText
import requests
from mailjet_rest import Client

# Load config
with open("config.properties", "r+b") as f:
    p = Properties()
    p.load(f, "utf-8")

    exchange_mode = p.properties.get("exchange.mode")
    smtp_server = p.properties.get("smtp.server")
    smtp_port = p.properties.get("smtp.port")
    smtp_user = p.properties.get("smtp.user")
    smtp_password = p.properties.get("smtp.password")
    mailjet_api_key = p.properties.get("mailjet.api.key")
    mailjet_api_secret = p.properties.get("mailjet.api.secret")
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
            filters = ["member@digitimes.com"]
            if not all(f in body for f in filters):
                send_notification(subject, body)


def send_notification(subject, body):
    mimet = MIMEText(body)
    mimet["Subject"] = subject
    mimet["From"] = smtp_user
    mimet["To"] = recipient
    msg = mimet.as_string()

    if exchange_mode == "smtp":
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

    if exchange_mode == "mailjet":
        api_key = mailjet_api_key
        api_secret = mailjet_api_secret
        mailjet = Client(auth=(api_key, api_secret), version='v3.1')
        data = {
            'Messages': [
                {
                    "From": {
                        "Email": smtp_user,
                        "Name": "Rojar Smith"
                              },
          "To": [
            {
              "Email": recipient,
              "Name": "Rojar"
            }
            ],
                    "Subject": subject,
                    "TextPart": "My first Mailjet email",
                    "HTMLPart": body,
                    "CustomID": "AppGettingStartedTest"
                }
            ]
        }
        result = mailjet.send.create(data=data)
        print(result.status_code)
        print(result.json())       


if __name__ == "__main__":
    outlook_handler = OutlookHandler()
    while True:
        pythoncom.PumpWaitingMessages()
    # print("test")
