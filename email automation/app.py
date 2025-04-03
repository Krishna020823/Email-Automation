from flask import Flask, request, render_template, jsonify
import pandas as pd
import win32com.client as win32
import re
import pythoncom
import threading
from bs4 import BeautifulSoup
import os
import tempfile

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Replace with a secure key

# Email validation pattern
email_pattern = re.compile(r"[^@]+@[^@]+")

# Global variable to store sent emails
sent_emails = []

def send_emails(df, rich_text_content, signature_text, attachments):
    pythoncom.CoInitialize()
    global sent_emails
    sent_emails = []

    try:
        outlook = win32.Dispatch("Outlook.Application")
        for index, row in df.iterrows():
            name = row["Name"]
            email = row["Email"]
            cc = row.get("CC", "")
            bcc = row.get("BCC", "")

            if not email_pattern.match(email):
                print(f"Invalid email address: {email}")
                continue

            print(f"Sending email to: {email}, CC: {cc}, BCC: {bcc}")

            mail = outlook.CreateItem(0)
            recipients = mail.Recipients
            recipients.Add(email)
            if cc:
                recipients.Add(cc)
            if bcc:
                recipients.Add(bcc)
            recipients.ResolveAll()

            soup = BeautifulSoup(rich_text_content, "html.parser")
            plain_text_content = soup.get_text()

            email_body = f"Dear {name},\n\n{plain_text_content}\n\n{signature_text}"
            mail.Subject = "Warning Notice"
            mail.Body = email_body

            if attachments:
                for attachment_name, attachment_data in attachments:
                    if attachment_name:  # Ensure attachment_name is not empty
                        temp_file_path = os.path.join(tempfile.gettempdir(), attachment_name)
                        with open(temp_file_path, 'wb') as temp_file:
                            temp_file.write(attachment_data)
                        mail.Attachments.Add(temp_file_path)
                        os.remove(temp_file_path)

            mail.Send()
            sent_emails.append(email)

        print("Emails sent successfully!")
        print("Sent emails:", sent_emails)
    finally:
        pythoncom.CoUninitialize()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            df = pd.read_excel(file)
            rich_text_content = request.form['rich_text_content']
            signature_text = request.form['signature_text']
            attachments = request.files.getlist('attachments')

            threading.Thread(target=send_emails, args=(df, rich_text_content, signature_text, attachments)).start()
            return jsonify({'total_emails': len(df)})

    return render_template('index.html', sent_emails=sent_emails)

@app.route('/start_sending_emails', methods=['POST'])
def start_sending_emails():
    file = request.files['file']
    if file:
        df = pd.read_excel(file)
        rich_text_content = request.form['rich_text_content']
        signature_text = request.form['signature_text']
        attachments = request.files.getlist('attachments')

        # Read attachment data
        attachment_data_list = []
        for attachment in attachments:
            attachment_data = attachment.read()
            attachment_data_list.append((attachment.filename, attachment_data))

        threading.Thread(target=send_emails, args=(df, rich_text_content, signature_text, attachment_data_list)).start()
        return jsonify({'total_emails': len(df)})

@app.route('/get_sent_emails')
def get_sent_emails():
    return jsonify(sent_emails)

if __name__ == '__main__':
    app.run(debug=True)