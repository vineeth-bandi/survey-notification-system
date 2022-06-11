import gspread
import httplib2
import os
import oauth2client
from oauth2client import file, client, tools
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apiclient import errors, discovery
from statistics import mean

SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Send Email'

gc = gspread.service_account()

sht1 = gc.open_by_key('1wo8OhfgWy7-ITOgnz0EQBdjQx0SBOIouM88QQKcRPpI')
worksheet = sht1.worksheet('Form Responses 1')

def compare_preference():
    actual_modes = worksheet.col_values(3)
    preferred_modes = worksheet.col_values(12)
    count = 0
    for i in range(1, len(actual_modes)):
        if actual_modes[i] == preferred_modes[i]:
            count += 1
    avg = count/(len(actual_modes)-1)
    if (avg > .5):
        print("notification sent.")

def check_satisfaction():
    modes = {
        'Arts and Crafts' : [], 
        'Book Club' : [], 
        'ESL Programs' : [], 
        'Speaker Events': [], 
        'Storytime' : []
        }

    for i in range(2, len(worksheet.col_values(1))+1):
        mode = worksheet.cell(i, 2).value
        val = worksheet.cell(i, 6).value
        modes[mode].append(int(val))

    for mode in modes:
        if len(modes[mode]) > 0 and mean(modes[mode]) < 1.5:
            print("notification sent.")

def check_social_connectivity():
    modes = {
        'Arts and Crafts' : [], 
        'Book Club' : [], 
        'ESL Programs' : [], 
        'Speaker Events': [], 
        'Storytime' : []
        }

    for i in range(2, len(worksheet.col_values(1))+1):
        mode = worksheet.cell(i, 2).value
        val = worksheet.cell(i, 7).value
        modes[mode].append(int(val))

    for mode in modes:
        if len(modes[mode]) > 0 and mean(modes[mode]) < 1.5:
            print("notification sent.")       

def check_engagement():
    modes = {
        'Arts and Crafts' : [], 
        'Book Club' : [], 
        'ESL Programs' : [], 
        'Speaker Events': [], 
        'Storytime' : []
        }

    for i in range(2, len(worksheet.col_values(1))+1):
        mode = worksheet.cell(i, 2).value
        val = worksheet.cell(i, 8).value
        modes[mode].append(int(val))

    for mode in modes:
        if len(modes[mode]) > 0 and mean(modes[mode]) < 1.5:
            print("notification sent.")        
    
def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'gmail-python-email-send.json')
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def SendMessage(sender, to, subject, msgHtml, msgPlain):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    message1 = CreateMessage(sender, to, subject, msgHtml, msgPlain)
    SendMessageInternal(service, "me", message1)

def SendMessageInternal(service, user_id, message):
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)

def CreateMessage(sender, to, subject, msgHtml, msgPlain):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to
    msg.attach(MIMEText(msgPlain, 'plain'))
    msg.attach(MIMEText(msgHtml, 'html'))
    raw = base64.urlsafe_b64encode(msg.as_bytes())
    raw = raw.decode()
    body = {'raw': raw}
    return body

def main():
    to = "ethicsgroup7@gmail.com"
    sender = "ethicsgroup7@gmail.com"
    subject = "Survey Warning Notification"
    msgHtml = "Hi<br/>Html Email"
    msgPlain = "Hi\nPlain Email"
    SendMessage(sender, to, subject, msgHtml, msgPlain)


if __name__ == '__main__':
    main()