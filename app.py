from __future__ import print_function
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
# import mediaFileUpload from googleapiclient

import pandas as pd
import datetime


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']


logFile = open("log.txt", "a+")
logFile.write("\nStarted at: " + str(datetime.datetime.now()))
# To list folders


def CheckFolder(service, FileName):
    page_token = None
    # response = service.files().list(q="mimeType = 'application/vnd.google-apps.spreadsheet'",
    # response = service.files().list(q="mimeType = 'application/vnd.google-apps.folder'",
    #                                 spaces='drive',
    #                                 fields='nextPageToken, files(id, name)',
    #                                 pageToken=page_token).execute()
    response = service.files().list(pageSize=1000, q="mimeType='application/vnd.google-apps.folder' and trashed = false",
                                    spaces='drive', fields='nextPageToken, files(id, name, mimeType, parents, sharingUser)', pageToken=page_token).execute()
    items = response.get('files', [])
    #     # Process change
    #     print('Found file: %s (%s)' % (file.get('name'), file.get('id')))
    # page_token = response.get('nextPageToken', None)
    # if page_token is None:
    #     break
    # for i in items:
    #     print(i['name'])
    if not items:
        print('No files found.')
        logFile.write("\nNo files found.")
        return None
    else:
        # print('Files:')
        for item in items:
            # print(item['name'])
            if(item['name'] == FileName):
                print(FileName + " is already there")
                logFile.write(FileName + " is already there")
                # print(item['name'])
                return item['id']


def get_excel_values():

    # read excel file
    df = pd.read_excel('kaizen942-attachments\Sample Sheet.xlsx')
    # print all values after 4th row
    # print(df.iloc[4:])
    return df.iloc[3:].values


def CheckFileDir(FileName, service):
    # page_token = None
    results = service.files().list(q="mimeType = 'application/vnd.google-apps.spreadsheet'",
                                   spaces='drive', fields="nextPageToken, files(id, name)", pageSize=400).execute()
    items = results.get('files', [])

    # print(len(items))
    # for i in items:
    if not items:
        logFile.write('\nNo files found.')
        print('No files found.')
        return None
    else:
        # print('Files:')
        for item in items:
            # print(item['name'])
            if(item['name'] == FileName):
                print(FileName + " is already there")
                logFile.write("\n"+FileName + " is already there")
                # print(item['name'])
                return item['id']


def CopyToFolder(folder, name, service):
    # Find Bata File
    File = CheckFileDir(name)
    print(File)
    # Find sector if not then create
    sector = CreateFolder(folder)
    newfile = {'name': name, 'parents': [sector]}
    service.files().copy(fileId=File, body=newfile).execute()
    print("Success copying file")
    logFile.write("\nSuccess copying file")


def CreateFolder(folder, service, parent=None):
    # Call the Drive v3 API
    # CheckFileDir(folder)
    print("folder/File is not there creating one...")
    logFile.write("\nfolder/File is not there creating one...")
    body = {
        'name': folder,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    if parent:
        body['parents'] = [parent]
    file = service.files().create(body=body,
                                  fields='id').execute()
    print('Folder ID: %s' % file.get('id'))
    logFile.write('\nFolder ID: %s' % file.get('id'))
    return file.get('id')
    # print(u'{0}'.format(item['name']))


def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)  # credentials.json download from drive API
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    excel_values = get_excel_values()
    # for i in range(len(values)):
    #     print(values[i][10])
    #     print(values[i][9])
    # extract filename from path
    filename = excel_values[0][10].split("\\")[-1].split(".")[0]
    print(filename)
    CheckFolder(service, excelvalues[0][9])
    file_metadata = {'name': filename,'parents' : [ folderId ]}

    media = MediaFileUpload(excel_values[0][10], mimetype='image/jpeg')
    file = service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()
    print('File ID: %s' % file.get('id'))
