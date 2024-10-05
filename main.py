import os
import base64
import re
from PIL import Image
import requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pytesseract
import fitz

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate_gmail():
    """Authenticate and create the Gmail API client."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def fetch_receipt_emails(service):
    """Fetch emails containing receipts and extract attachments."""
    results = service.users().messages().list(userId='me', q='receipt OR invoice').execute()
    messages = results.get('messages', [])
    receipt_details = []

    for message in messages:
        msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
        payload = msg['payload']
        headers = payload['headers']

        for header in headers:
            if header['name'] == 'Subject':
                subject = header['value']
                print(f'Processing email with subject: {subject}')

        if 'parts' in payload:  # Handle multipart messages
            for part in payload['parts']:
                if part['filename']:  # Check if there's an attachment
                    if part['mimeType'] in ['application/pdf', 'image/jpeg', 'image/png']:
                        attachment_id = part['body']['attachmentId']
                        attachment = service.users().messages().attachments().get(userId='me', messageId=message['id'], id=attachment_id).execute()
                        data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                        # Save the attachment locally
                        file_path = os.path.join('attachments', part['filename'])
                        with open(file_path, 'wb') as f:
                            f.write(data)
                        print(f'Saved attachment: {file_path}')
                        receipt_details.append({'subject': subject, 'attachment': file_path})

    return receipt_details

def ocr_receipt(file_path):
    extracted_text = ""
    
    # Check if the file is a PDF
    if file_path.endswith('.pdf'):
        # Open the PDF file
        pdf_document = fitz.open(file_path)
        
        # Iterate through the pages
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            extracted_text += page.get_text()  # Extract text from each page
            
        pdf_document.close()
    else:
        # If it's an image, use the original OCR method
        image = Image.open(file_path)
        # Convert the image to a format suitable for OCR (if needed)
        extracted_text = pytesseract.image_to_string(image)
    
    return extracted_text

def process_receipts(receipts):
    """Process each receipt document for OCR."""
    for receipt in receipts:
        file_path = receipt['attachment']
        extracted_text = ocr_receipt(file_path)

        if extracted_text:
            print(f"Processed: {receipt['subject']}")
            print("Extracted Text:", extracted_text)  # Handle extracted text as needed

def main():
    service = authenticate_gmail()
    receipts = fetch_receipt_emails(service)

    # Process each receipt with OCR
    process_receipts(receipts)

if __name__ == '__main__':
    main()
