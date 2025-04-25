from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import json

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents',
          'https://www.googleapis.com/auth/drive']

def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def create_meal_plan_document():
    creds = get_credentials()
    
    # Create Google Docs and Drive services
    docs_service = build('docs', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    # Create a new document
    document = docs_service.documents().create(body={
        'title': 'Meal Plan Template'
    }).execute()
    doc_id = document.get('documentId')

    # Create dropdown options
    breakfast_items = [
        'بيض مسلوق',
        'شوفان',
        'توست اسمر',
        'جبنة قليلة الدسم',
        'لبنة'
    ]

    lunch_items = [
        'صدر دجاج مشوي',
        'سمك مشوي',
        'لحم بقر مشوي',
        'أرز بني',
        'بطاطا مسلوقة'
    ]

    dinner_items = [
        'سلطة خضراء',
        'شوربة خضار',
        'خضار مشوية'
    ]

    # Create a form in the document
    form_metadata = {
        'title': 'Meal Plan Form',
        'documentType': 'FORM',
        'fields': []
    }

    # Add days of the week with dropdown fields
    days = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت']
    
    requests = []
    
    # Add title
    requests.append({
        'insertText': {
            'location': {
                'index': 1
            },
            'text': 'خطة الوجبات الأسبوعية\n\n'
        }
    })

    current_index = 30  # Starting index after title

    for day in days:
        # Add day header
        requests.extend([
            {
                'insertText': {
                    'location': {
                        'index': current_index
                    },
                    'text': f'{day}\n'
                }
            },
            # Add meal type labels and dropdowns
            {
                'insertText': {
                    'location': {
                        'index': current_index + len(day) + 1
                    },
                    'text': 'الفطور: [dropdown]\n'
                }
            },
            {
                'insertText': {
                    'location': {
                        'index': current_index + len(day) + 20
                    },
                    'text': 'الغداء: [dropdown]\n'
                }
            },
            {
                'insertText': {
                    'location': {
                        'index': current_index + len(day) + 40
                    },
                    'text': 'العشاء: [dropdown]\n\n'
                }
            }
        ])
        
        current_index += 80  # Increment index for next day

    # Execute the requests
    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={'requests': requests}
    ).execute()

    # Create a form and link it to the document
    form_metadata = {
        'linkedDocumentId': doc_id,
        'formFields': []
    }

    # Add form fields for each dropdown
    for day in days:
        form_metadata['formFields'].extend([
            {
                'fieldId': f'{day}_breakfast',
                'fieldType': 'DROP_DOWN',
                'title': f'الفطور - {day}',
                'options': breakfast_items
            },
            {
                'fieldId': f'{day}_lunch',
                'fieldType': 'DROP_DOWN',
                'title': f'الغداء - {day}',
                'options': lunch_items
            },
            {
                'fieldId': f'{day}_dinner',
                'fieldType': 'DROP_DOWN',
                'title': f'العشاء - {day}',
                'options': dinner_items
            }
        ])

    # Create form
    form = drive_service.files().create(
        body=form_metadata,
        fields='id'
    ).execute()

    print(f'Created document with ID: {doc_id}')
    print(f'Created form with ID: {form.get("id")}')
    
    # Set document permissions to anyone with link can view
    drive_service.permissions().create(
        fileId=doc_id,
        body={
            'type': 'anyone',
            'role': 'reader'
        }
    ).execute()

    return doc_id

if __name__ == '__main__':
    doc_id = create_meal_plan_document()
    print(f'View your document at: https://docs.google.com/document/d/{doc_id}')