"""
Step 1: Set Up a Google Cloud Project
    Go to the Google Cloud Console.
    Create a new project or select an existing one.
    Navigate to APIs & Services > Credentials.
    Click on Create Credentials and select OAuth 2.0 Client IDs.
    Configure the consent screen, fill in the necessary details.
    Choose Desktop app as the application type.
    Download the OAuth 2.0 client credentials JSON file.

pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

pip install pandas openpyxl

pip install flask flask-wtf

"""

from flask import Flask, render_template, request, redirect, flash, url_for
import os
import base64
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "supersecretkey"  # Replace with a secure key for session management
UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def get_credentials():
    """Gets OAuth2 credentials and saves them to token.json"""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds


def create_message(sender_name, sender_email, to, subject, message_text):
    """Creates a MIME message for email with sender's name"""
    from_email = f"{sender_name} <{sender_email}>"
    message = MIMEText(message_text)
    message["to"] = to
    message["from"] = from_email
    message["subject"] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {"raw": raw}


def send_message(service, user_id, message):
    """Sends the email message via Gmail API"""
    try:
        message = (
            service.users().messages().send(userId=user_id, body=message).execute()
        )
        return message["id"]
    except Exception as error:
        print(f"An error occurred: {error}")
        return None


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Handle file upload
        file = request.files["file"]
        sender_name = request.form["sender_name"]
        sender_email = request.form["sender_email"]
        subject = request.form["subject"]
        message_body = request.form["message_body"]

        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)

            # Read recipients from Excel file
            try:
                df = pd.read_excel(filepath)
                recipients = df[["Name", "Email"]].dropna()

                # Send emails
                creds = get_credentials()
                service = build("gmail", "v1", credentials=creds)

                for _, row in recipients.iterrows():
                    name = row["Name"]
                    recipient = row["Email"]
                    personalized_message = message_body.replace("{name}", name)
                    email_message = create_message(
                        sender_name,
                        sender_email,
                        recipient,
                        subject,
                        personalized_message,
                    )
                    send_message(service, sender_email, email_message)

                flash("Emails sent successfully!", "success")
            except Exception as e:
                flash(f"An error occurred: {e}", "danger")
        else:
            flash("Please upload a valid Excel file.", "warning")
        return redirect(url_for("index"))

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
