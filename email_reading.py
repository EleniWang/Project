import xlsxwriter
import pdfplumber
import os
import pickle
import re
import os.path
from apiclient import errors
import email
from email.mime.text import MIMEText
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
from bs4 import BeautifulSoup
from lxml import etree

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# Get all the ids of unread message
def search_message_unread(service, user_id):
    try:
        list_ids = []
        search_ids = service.users().messages().list(userId='me',labelIds=['INBOX'],q="is:unread").execute() 
        try:
            ids = search_ids['messages']

        except KeyError:
            print("WARNING: the search queried returned 0 results")
            print("returning an empty string")
            return ""

        if len(ids)>1:
            for msg_id in ids:
                list_ids.append(msg_id['id'])
            return(list_ids)

        elif len(ids)==1:
            list_ids.append(ids[0]['id'])
            return list_ids
        else:
            pass
        
    except (errors.HttpError, error):
        print("An error occured: %s") % error

# get all the unread messages
def get_message_unread(service, user_id, msg_id):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id,format='raw').execute()
        unread = service.users().messages().list(userId='me',labelIds=['INBOX'],q="is:unread").execute() 
        messages = unread.get('messages',[])
        for message in messages:
            service.users().messages().modify(userId = 'me',id=message['id'],body = {'removeLabelIds':['UNREAD']}).execute()
    except Exception:
        print("An error occured: %s") % error

# get all the ids of messages sending through Domain.com
def search_message_domain(service, user_id, search_string):
    """
    Search the inbox for emails using standard gmail search parameters
    and return a list of email IDs for each result
    PARAMS:
        service: the google api service object already instantiated
        user_id: user id for google api service ('me' works here if
        already authenticated)
        search_string: search operators you can use with Gmail
        (see https://support.google.com/mail/answer/7190?hl=en for a list)
    RETURNS:
        List containing email IDs of search query
    """
    try:
        # initiate the list for returning
        list_ids = []

        # get the id of all messages that are in the search string
        # search_ids = service.users().messages().list(userId=user_id, q=search_string).execute()
        search_ids = service.users().messages().list(userId='me',labelIds=['INBOX'],q="is:unread").execute() 
        search_ids_1 = service.users().messages().list(userId=user_id, q=search_string).execute()
        
        # if there were no results, print warning and return empty string
        try:
            ids_1 = search_ids['messages']
            ids_2 = search_ids_1['messages']
            ids = []
            for i in ids_1:
                if i in ids_2:
                    ids.append(i)

        except KeyError:
            print("WARNING: the search queried returned 0 results")
            print("returning an empty string")
            return ""

        if len(ids)>1:
            for msg_id in ids:
                list_ids.append(msg_id['id'])
            return(list_ids)

        elif len(ids) == 1:
            list_ids.append(ids[0]['id'])
            return list_ids
        else:
            pass
        
    except (errors.HttpError, error):
        print("An error occured: %s") % error

# get all the messages sending by Domain.com
def get_message_domain(service, user_id, msg_id):
    """
    Search the inbox for specific message by ID and return it back as a 
    clean string. String may contain Python escape characters for newline
    and return line. 
    
    PARAMS
        service: the google api service object already instantiated
        user_id: user id for google api service ('me' works here if
        already authenticated)
        msg_id: the unique id of the email you need
    RETURNS
        A string of encoded text containing the message body
    """
    try:
        # grab the message instance
        message = service.users().messages().get(userId=user_id, id=msg_id,format='raw').execute()

        # decode the raw string, ASCII works pretty well here
        msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))

        # grab the string from the byte object
        mime_msg = email.message_from_bytes(msg_str)

        # check if the content is multipart (it usually is)
        content_type = mime_msg.get_content_maintype()
        if content_type == 'multipart':
            # there will usually be 2 parts the first will be the body in text
            # the second will be the text in html
            parts = mime_msg.get_payload()

            # return the encoded text
            final_content = parts[0].get_payload()
            #print(final_content)

        elif content_type == 'text':
            parts = mime_msg.get_payload()
            soup_1 = BeautifulSoup(parts,'lxml')
            table = soup_1.table
            tr_arr = table.find_all("tr")
            f_l = []
            person = []
            for tr in tr_arr:
                tds = tr.find_all('td')
                file = tds[0].get_text()
                final = file.replace('\r','').replace('\n','').replace('\t','').replace('=','')
                if final.isspace() == False:
                    if final != '</span>':
                        f_l.append(final)
            print(f_l[0])
            print( )
            result = re.findall(".*property at(.*)ref.*",f_l[0])
            l = result[0].split()
            print(l)
            address = l[0].replace('20',' ')+' '
            for i in range(1,len(l)-1):
                address += l[i]
                address += ' '
            address += l[-1][:-5]
            person.append(address)
            print(address)
            result_1 = re.findall(".*From:(.*)Email.*",f_l[0])
            name = ''
            for i in result_1:
                name += i
            person.append(name)
            result_2 = re.findall(".*Phone:(.*)Message.*",f_l[0])
            number = ''
            for i in result_2:
                number += i
            person.append(number)
            result_3 = re.findall(".*Message:(.*)Security Policy.*",f_l[0])
            message = ''
            for i in result_3:
                message += i
            person.append(message.replace('-',''))
            result_4 = re.search(r'(%s.*?%s)'%('Email:','Phone'),f_l[0]).group(1)
            e_address = result_4[6:-5]
            person.append(e_address)
            person.append('Domain/REA')
            return person
        else:
            return ""
            print("\nMessage is not text or multipart, returned an empty string")
    # unsure why the usual exception doesn't work in this case, but 
    # having a standard Exception seems to do the trick
    except Exception:
        print("An error occured: %s") % error

# get all the ids of messages from realestate.com
def search_message_realestate(service, user_id, search_string):
    try:
        # initiate the list for returning
        list_ids = []

        # get the id of all messages that are in the search string
        search_ids = service.users().messages().list(userId='me',labelIds=['INBOX'],q="is:unread").execute() 
        search_ids_1 = service.users().messages().list(userId=user_id, q=search_string).execute()
        
        # if there were no results, print warning and return empty string
        try:
            ids_1 = search_ids['messages']
            ids_2 = search_ids_1['messages']
            ids = []
            for i in ids_1:
                if i in ids_2:
                    ids.append(i)

        except KeyError:
            print("WARNING: the search queried returned 0 results")
            print("returning an empty string")
            return ""

        if len(ids)>1:
            for msg_id in ids:
                list_ids.append(msg_id['id'])
            return(list_ids)

        elif len(ids)==1:
            list_ids.append(ids[0]['id'])
            return list_ids
        else:
            pass
        
    except (errors.HttpError, error):
        print("An error occured: %s") % error

# get all the messages from realestate.com        
def get_message_realestate(service, user_id, msg_id):
    """
    Search the inbox for specific message by ID and return it back as a 
    clean string. String may contain Python escape characters for newline
    and return line. 
    
    PARAMS
        service: the google api service object already instantiated
        user_id: user id for google api service ('me' works here if
        already authenticated)
        msg_id: the unique id of the email you need
    RETURNS
        A string of encoded text containing the message body
    """
    try:
        # grab the message instance
        message = service.users().messages().get(userId=user_id, id=msg_id,format='raw').execute()

        # decode the raw string, ASCII works pretty well here
        msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))

        # grab the string from the byte object
        mime_msg = email.message_from_bytes(msg_str)

        # check if the content is multipart (it usually is)
        content_type = mime_msg.get_content_maintype()
        if content_type == 'multipart':
            # there will usually be 2 parts the first will be the body in text
            # the second will be the text in html
            parts = mime_msg.get_payload()

            # return the encoded text
            final_content = parts[0].get_payload()
            return final_content

        elif content_type == 'text':
            person = []
            parts = mime_msg.get_payload()
            soup_1 = BeautifulSoup(parts,'lxml')
            a = soup_1.get_text().replace('\n','')
            result_1 = re.findall(".*address:(.*)Property URL.*",a)
            address = result_1[0]
            person.append(address)
            result_1 = re.findall(".*Name:(.*)Email.*",a)
            name = result_1[0]
            person.append(name)
            result_1 = re.findall(".*Phone:(.*)About me.*",a)
            number = result_1[0]
            person.append(number)
            result_1 = re.findall(".*About me:(.*)You can.*",a)
            messages = result_1[0].replace('Comments:','')
            person.append(messages)
            result_1 = re.findall(".*Email:(.*)Phone.*",a)
            e_address = result_1[0]
            person.append(e_address)
            person.append('Domain/REA')
            return person

        else:
            return ""
            print("\nMessage is not text or multipart, returned an empty string")
    # unsure why the usual exception doesn't work in this case, but 
    # having a standard Exception seems to do the trick
    except Exception:
        print("An error occured: %s") % error
        
def get_service():
    """
    Authenticate the google api client and return the service object 
    to make further calls
    PARAMS
        None
    RETURNS
        service api object from gmail for making calls
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
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)


    service = build('gmail', 'v1', credentials=creds)

    return service

# saving to the xlsx
xl = xlsxwriter.Workbook(r'C:\Users\Eleni\Desktop\new_lead.xlsx')
sheet = xl.add_worksheet('sheet1')
sheet.write_string("A1","address")
sheet.write_string("B1","name")
sheet.write_string("C1","number")
sheet.write_string("D1","message")
sheet.write_string("E1","email")
sheet.write_string("F1","source")
x=0
total =[]
message_id = search_message_realestate(service = get_service(),user_id = 'me',search_string = "realestate")
if message_id != None:
    for i in message_id:
        a = get_message_realestate(service = get_service(),user_id = 'me',msg_id = i)
        if a != None:
            total.append(a)

message_id = search_message_domain(service = get_service(),user_id = 'me',search_string = "The enquiry has been sent from a user on www.domain.com.au")
if message_id != None:
    for i in message_id:
        a = get_message_domain(service = get_service(),user_id = 'me',msg_id = i)
        if a != None:
            total.append(a)
for i in range(len(total)):
    x += 1
    y = x+1
    sheet.write_string("A%d" % y, "%s" % total[i][0])
    sheet.write_string("B%d" % y, "%s" % total[i][1])
    sheet.write_string("C%d" % y, "%s" % total[i][2])
    sheet.write_string("D%d" % y, "%s" % total[i][3])
    sheet.write_string("E%d" % y, "%s" % total[i][4])
    sheet.write_string("F%d" % y, "%s" % total[i][5])
xl.close()
message_id = search_message_unread(service = get_service(),user_id = 'me')
if message_id != None:
    for i in message_id:
        a = get_message_unread(service = get_service(),user_id = 'me',msg_id = i)



# changing the format from xlsx to csv
import pandas as pd
def xlsx_to_csv_pd():
    data_xls = pd.read_excel(os.getcwd() + '\\' + 'Desktop' + '\\' + 'new_lead.xlsx', index_col=0)
    data_xls.to_csv(os.getcwd() + '\\' + 'Desktop' + '\\' + '\\' +'new_lead.csv', encoding='utf-8')
if __name__ == '__main__':
    xlsx_to_csv_pd()






