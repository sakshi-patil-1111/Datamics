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
import datefinder
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Path to Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update path as needed

# Gmail SCOPES
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Email notification settings
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = 'preetkaurpawar8@gmail.com'
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')  # Ensure it's set in the environment
EMAIL_RECIPIENTS = ['datamicsbombay@gmail.com']

# Google Form URL and input field names
GOOGLE_FORM_URL = 'https://docs.google.com/forms/d/e/1FAIpQLScKLeWRp0zpv08wpHCWyoMWl8-TVEeVVcyoRKec_FOm07ttPw/formResponse'
FORM_FIELDS = {
    'receipt_date': 'entry.23770198',          # Entry ID for the "Receipt Date" field
    'receipt_number': 'entry.855132866',       # Entry ID for the "Receipt Number" field
    'vendor_name': 'entry.793740564',          # Entry ID for the "Vendor Name" field
    'total_amount': 'entry.963910700',         # Entry ID for the "Total Amount" field
    'items_purchased': 'entry.1797858785'      # Entry ID for the "Items Purchased" field
}

def send_email(subject, body, recipients):
    """Send an email notification."""
    try:
        if not subject or not body:
            logger.error("Email subject or body is empty. Subject: %s, Body: %s", subject, body)
            return
        
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        text = msg.as_string()
        server.sendmail(EMAIL_ADDRESS, recipients, text)
        server.quit()
        logger.info("Email notification sent successfully.")
    except Exception as e:
        logger.error(f"Failed to send email: {e}")


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
    """Fetch emails containing food or petrol receipts and extract attachments."""
    results = service.users().messages().list(userId='me', q='food OR petrol receipt OR invoice').execute()
    messages = results.get('messages', [])
    receipt_details = []

    for message in messages:
        msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
        payload = msg['payload']
        headers = payload['headers']

        subject = None
        for header in headers:
            if header['name'] == 'Subject':
                subject = header['value']
                logger.info(f'Processing email with subject: {subject}')

        if 'parts' in payload:  # Handle multipart messages
            for part in payload['parts']:
                if part.get('filename'):  # Check if there's an attachment
                    if part['mimeType'] in ['application/pdf', 'image/jpeg', 'image/png']:
                        attachment_id = part['body']['attachmentId']
                        attachment = service.users().messages().attachments().get(
                            userId='me',
                            messageId=message['id'],
                            id=attachment_id
                        ).execute()
                        data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))

                        # Save the attachment locally
                        os.makedirs('attachments', exist_ok=True)
                        file_path = os.path.join('attachments', part['filename'])
                        with open(file_path, 'wb') as f:
                            f.write(data)
                        logger.info(f'Saved attachment: {file_path}')
                        receipt_details.append({'subject': subject, 'attachment': file_path})

    return receipt_details

def ocr_receipt(file_path):
    """Perform OCR on the receipt attachment to extract text."""
    extracted_text = ""
    
    if file_path.lower().endswith('.pdf'):
        try:
            pdf_document = fitz.open(file_path)
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                extracted_text += page.get_text()
            pdf_document.close()
        except Exception as e:
            logger.error(f"Failed to open PDF {file_path}: {e}")
    else:
        try:
            image = Image.open(file_path)
            extracted_text = pytesseract.image_to_string(image)
        except Exception as e:
            logger.error(f"Failed to perform OCR on image {file_path}: {e}")
    
    return extracted_text

def extract_receipt_data(extracted_text):
    """Extract relevant data from the OCR-extracted text."""
    data = {}

    # Extract date
    matches = list(datefinder.find_dates(extracted_text, source=True, index=True))
    if matches:
        data['date'] = min([match[0] for match in matches]).strftime('%Y-%m-%d')

    # Extract amounts
    amount_patterns = [r'(\$|USD)\s?([0-9]+(?:\.[0-9]{2})?)', r'([0-9]+(?:\.[0-9]{2})?)\s?(USD|\$)']
    amounts = [float(match[1]) for pattern in amount_patterns for match in re.findall(pattern, extracted_text)]
    data['amount'] = max(amounts) if amounts else None

    # Extract receipt number
    receipt_number_pattern = r'Receipt\s*[#:]\s*([0-9]+)'
    receipt_number_match = re.search(receipt_number_pattern, extracted_text, re.IGNORECASE)
    data['receipt_number'] = receipt_number_match.group(1) if receipt_number_match else None

    # Extract vendor (assumed first non-date, non-amount line)
    lines = extracted_text.split('\n')
    vendor = None
    for line in lines:
        if (data.get('date') and data['date'] in line) or (data.get('amount') and str(data['amount']) in line):
            continue
        if line.strip():
            vendor = line.strip()
            break
    data['vendor'] = vendor

    # Extract items (for food receipts)
    if 'food' in extracted_text.lower():
        data['items'] = "; ".join([line.strip() for line in lines if re.search(r'\b(Item|Qty|Amount)\b', line, re.IGNORECASE)])

    return data

def submit_to_google_form(data):
    """Submit extracted data to Google Form."""
    try:
        form_data = {
            FORM_FIELDS['receipt_date']: data.get('date', ''),
            FORM_FIELDS['receipt_number']: data.get('receipt_number', ''),
            FORM_FIELDS['vendor_name']: data.get('vendor', ''),
            FORM_FIELDS['total_amount']: data.get('amount', ''),
            FORM_FIELDS['items_purchased']: data.get('items', 'N/A')
        }
        
        # Log the form data before submission
        logger.info("Form Data: %s", form_data)

        response = requests.post(GOOGLE_FORM_URL, data=form_data)
        logger.info("Response Status Code: %s", response.status_code)
        # logger.info("Response Content: %s", response.text)

        if response.status_code == 200:
            logger.info("Data successfully submitted to Google Form.")
        else:
            logger.error(f"Failed to submit to Google Form: {response.status_code}")
    except Exception as e:
        logger.error(f"Error submitting data to Google Form: {e}")


def main():
    """Main function to orchestrate the automation."""
    service = authenticate_gmail()
    receipt_emails = fetch_receipt_emails(service)

    for email in receipt_emails:
        file_path = email['attachment']
        extracted_text = ocr_receipt(file_path)
        receipt_data = extract_receipt_data(extracted_text)
        
        if receipt_data.get('amount') and receipt_data.get('vendor'):
            submit_to_google_form(receipt_data)
            # Send email notification
            email_body = f"""
            Vendor: {receipt_data.get('vendor')}
            Date: {receipt_data.get('date')}
            Amount: {receipt_data.get('amount')}
            Items: {receipt_data.get('items', 'N/A')}
            """
            send_email(f"Receipt Processed: {receipt_data['vendor']}", email_body, EMAIL_RECIPIENTS)

if __name__ == '__main__':
    main()
