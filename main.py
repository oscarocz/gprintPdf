import os
import base64

from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly", "https://www.googleapis.com/auth/gmail.modify"]

# Authenticate Gmail User
def gmail_authenticate():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    elif os.path.exists("credentials.json") == False:
        print("Not token.json or credentials.json found")
        return False
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or creds.token_state == 'INVALID':
        if creds and (creds.token_state != 'FRESH') and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)

# Get Email Address
def gmail_readEmailAddress(service):
    result = service.users().getProfile(userId='me').execute()
    print("Email Address:", result['emailAddress'])

# Read unread Messages with Pdfs Attached
def gmail_readUnreadMessagesWithPdfs(service):
    query = 'is:unread AND has:attachment filename:pdf'
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q=query).execute()

    if results['resultSizeEstimate'] == 0:
        print("No email with bills")
        return 0
    
    # Set Messages with Pfds as Read
    msgIds = [msg['id'] for msg in results['messages']]
    service.users().messages().batchModify(userId='me', body={'removeLabelIds': ['UNREAD'], 'ids':msgIds}).execute()

    # Get PDFs Names & IDs
    pdfIdsMsg = {} 
    for msgId in msgIds:
        pdfIdsInMsg = []
        parts = service.users().messages().get(userId='me', id=msgId, format='full').execute()['payload']['parts']
        for part in parts:
            if part.get('mimeType') == 'application/pdf':
                fileName = part.get('filename')
                attachmentId = part['body'].get('attachmentId')
                pdfIdsInMsg.append(tuple([fileName, attachmentId]))
        pdfIdsMsg.update({msgId:pdfIdsInMsg})
    return pdfIdsMsg

# Create Different Download Folders each time the Script is Executed
def createDownloadFolder():
    currentDate = datetime.today().strftime('%Y-%m-%d')
    savePath = ""
    index = ""

    while True:
        try:
            savePath = os.path.join('downloads', currentDate + index)
            os.makedirs(savePath)
            break
        except WindowsError:
            if index:
                # Append 1 to number in brackets
                index = '('+str(int(index[1:-1])+1)+')'
            else:
                index = '(1)'
            pass
    return savePath

# Download Pdfs
def gmail_downloadPdfs(service, pdfIdsMsg: dict, folderPath: str):
    numPdfs = 0

    for msgId in pdfIdsMsg:
        numPdfs += len(pdfIdsMsg[msgId])
        for pdfId in pdfIdsMsg[msgId]:
            fileName = pdfId[0]
            attachmentId = pdfId[1]
            attachment =  service.users().messages().attachments().get(userId='me', messageId=msgId, id=attachmentId).execute()
            fileData = attachment.get('data')
            fileData = base64.urlsafe_b64decode(fileData)
            savePath = os.path.join(folderPath, fileName)
            with open(savePath, 'wb') as f:
                f.write(fileData)
    print("Number of PDFs downloaded:", numPdfs)

# Print folder Pdfs
def print_pdf(folderPath: str):
    count = 0
    for file in os.listdir(folderPath):
        if file.lower().endswith('.pdf') == True:
            count += 1
            print("Print:", file)
            filePath = os.path.join(folderPath, file)
            os.startfile(filePath, "print")
    print("Files printed:", count)

# Read unread Messages with Pdfs Attached
def gmail_checkNumberUnreadMessages(service):
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q='is:unread').execute()
    if results['resultSizeEstimate'] != 0:
        print('Inbox Unread Mails:', results['resultSizeEstimate'])

    results = service.users().messages().list(userId='me', labelIds=['SPAM'], q='is:unread').execute()
    if results['resultSizeEstimate'] != 0:
        print('Spam Unread Mails:', results['resultSizeEstimate'])

def main():
    print("Execute Print Email")

    try:
        # Get access to gmail service from credentials
        service = gmail_authenticate()
        if service == False:
            return
        
        gmail_readEmailAddress(service)

        # Check for unread Pdfs
        pdfIdsMsg = gmail_readUnreadMessagesWithPdfs(service)
        if pdfIdsMsg != 0:
            # Download Pdfs and print them
            folderPath = createDownloadFolder()
            gmail_downloadPdfs(service, pdfIdsMsg, folderPath)
            print_pdf(folderPath)

        gmail_checkNumberUnreadMessages(service)

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f"An error occurred: {error}")


if __name__ == '__main__':
    main()