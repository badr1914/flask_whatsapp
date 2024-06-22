from flask import Flask, request, send_file
import pandas as pd
from openpyxl import load_workbook, Workbook
from twilio.twiml.messaging_response import MessagingResponse
import os

app = Flask(__name__)

@app.route('/')
def index():
    return "Welcome to the Flask WhatsApp Integration!"

@app.route('/webhook', methods=['POST'])
def webhook():
    print(f"Received Content-Type: {request.content_type}")  # Debug statement
    if request.content_type == 'application/x-www-form-urlencoded':
        from_number = request.form['From']
        body = request.form['Body']
        print(f"Received message from {from_number}: {body}")  # Debug statement
        save_to_excel(from_number, body)
        resp = MessagingResponse()
        resp.message("Message received and saved.")
        return str(resp)
    else:
        return "Unsupported Media Type", 415

@app.route('/download', methods=['GET'])
def download_file():
    file_path = 'messages.xlsx'
    print(f"Checking if {file_path} exists...")  # Debug statement
    if os.path.exists(file_path):
        print(f"{file_path} found. Sending file...")  # Debug statement
        return send_file(file_path, as_attachment=True)
    else:
        print(f"{file_path} not found.")  # Debug statement
        return "No messages saved yet.", 404

def save_to_excel(sender, message):
    file_path = 'messages.xlsx'
    sheet_name = 'Sheet1'
    print(f"Saving message from {sender}: {message} to {file_path}")  # Debug statement
    try:
        book = load_workbook(file_path)
    except FileNotFoundError:
        print(f"{file_path} not found. Creating new workbook.")  # Debug statement
        book = Workbook()
        book.save(file_path)
        book = load_workbook(file_path)
    writer = pd.ExcelWriter(file_path, engine='openpyxl')
    writer.book = book
    df = pd.DataFrame([[sender, message]], columns=['Sender', 'Message'])
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=writer.sheets[sheet_name].max_row if sheet_name in writer.sheets else 1)
    writer.save()
    print(f"Message saved to {file_path}")  # Debug statement

if __name__ == '__main__':
    app.run(debug=True)
