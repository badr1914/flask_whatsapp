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
    print(f"Received request: {request}")  # Debug statement
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
        print("Unsupported Media Type")  # Debug statement
        return "Unsupported Media Type", 415

@app.route('/download', methods=['GET'])
def download_file():
    file_path = 'messages.xlsx'
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "No messages saved yet.", 404

def save_to_excel(sender, message):
    file_path = 'messages.xlsx'
    sheet_name = 'Sheet1'
    print(f"Saving message from {sender}: {message} to {file_path}")  # Debug statement
    
    if not os.path.exists(file_path):
        print(f"{file_path} not found. Creating new workbook.")  # Debug statement
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        headers = ['Sender', 'Message']
        ws.append(headers)
        wb.save(file_path)
    
    book = load_workbook(file_path)
    writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a')
    writer.book = book
    writer.sheets = {ws.title: ws for ws in book.worksheets}

    df = pd.DataFrame([[sender, message]], columns=['Sender', 'Message'])
    
    if sheet_name in writer.sheets:
        startrow = writer.sheets[sheet_name].max_row
    else:
        startrow = 0

    print(f"Writing to row: {startrow}")  # Debug statement
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=startrow)
    writer.save()
    writer.close()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
