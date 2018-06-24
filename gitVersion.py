#! python3

# The first function opens the input inbox,
# analyzes all of the messages in that inbox
# and outputs a list that includes:
# * reciever and sender addresses 
# * subject line
# * sent and recieve times
# * response time

# Requires imapclient, imaplib, pyzmail, openpyxl, pprint, re, datetime

def getAddAndSub(username,pword,hostAddress):

    # Increases the max allowed size of program to 10,000,000
    # bytes, much less likely to get this error:
    # imaplib.error: got more than 10000 bytes
    import imaplib
    imaplib._MAXLINE = 10000000
    
    import re
    responseDateRegex = re.compile(r'Date: (\w\w\w, \d\d \w\w\w \d\d\d\d \d\d:\d\d:\d\d (-|\+)\d\d\d\d)')
    originalDateRegex = re.compile(r'Sent: ((Mon|Tues|Wednes|Thurs|Fri|Satur|Sun)day, (January|February|March|April|May|June|July|August|September|October|November|December) (\d\d|\d), \d\d\d\d (\d\d|\d):\d\d (A|P)M)')

    import imapclient
    from datetime import datetime
    with imapclient.IMAPClient(host=hostAddress) as client:

        # Login and access the inbox.
        try:
            client.login(username,pword)
        except imaplib.IMAP4.error:
            print('The password is incorrect. Please run the program again.')
            input()
            raise SystemExit
        
#        except TimeoutError:
#            print('''The connection attempt failed because the connected
#host has failed to respond.
#Please check the name of your IMAP mail server and try again.''')
#            input()
#            raise SystemExit
                  
        client.select_folder('INBOX',readonly=True)

        # Aquire the message #'s of all messages in the inbox.
        messages = client.search(['ALL'])
        length = len(messages)

        print('Your password was correct. Looking through ' + str(length) + ' emails...')
        info = []
        import pyzmail

        for i in range(length):   
            info.append(['','','','','','',''])
            rawMessages = client.fetch(messages[i], ['BODY[]'])
            legibleMessage = pyzmail.PyzMessage.factory(rawMessages[messages[i]][b'BODY[]'])

            exitDate = responseDateRegex.search(str(legibleMessage))           
            recieveDate = originalDateRegex.search(str(rawMessages))
            
            info[i][0] = messages[i]
            info[i][1] = legibleMessage.get_address('from')
            info[i][2] = legibleMessage.get_address('to')
            info[i][3] = legibleMessage.get_subject()
            info[i][5] = datetime.strptime(exitDate.group(1), '%a, %d %b %Y %H:%M:%S %z')

            if recieveDate != None:
                info[i][4] = datetime.strptime(str(recieveDate.group(1)) + ' -0400', '%A, %B %d, %Y %I:%M %p %z')
                info[i][6] = info[i][5] - info[i][4]
            
            if (i + 1) % 10 == 0 and (i + 1) != length:
                print(str(i + 1) + '/' + str(length) + ' messages analyzed.')
        client.logout()
        print('All messages analyzed.')
    return info
    # Output is the unique IDs (UIDs) subject, sender
    # and reciever of each email, and times that the
    # quotes were recieved and responded to.

def createNewXlsx():
    import os

    print('''
Your new excel file with the email data will be
created in the same folder as this python script.
Make sure this .xlsx doesn't share a name as 
another .xlsx in the same folder.

What do you want to name this file?''')
    fileName = input()
    print('\nLets try that.')
    path = str(os.getcwd())
    print('The current file path is "' + path + '"')

    import openpyxl
    wb = openpyxl.Workbook()
    sheet = wb['Sheet']
    sheet['A1'] = 'Message UID'
    sheet['B1'] = 'Sender'
    sheet['C1'] = 'Recipient (Client)'
    sheet['D1'] = 'Subject Line'
    sheet['E1'] = 'Quote Request Date/Time'
    sheet['F1'] = 'Quote Response Date/Time'
    sheet['G1'] = 'Turnaround [HR:MIN:SEC]'
    
    try:
        wb.save(path + '\\' + fileName + '.xlsx')
        print('\nThe file ' + fileName + '.xslx has been saved in the above folder.')
    except PermissionError:
        print('''
***Another file with the same name is open!***
Close the file and try again if you are okay with saving over it.''')
        createNewXlsx()
    fileAndPath = [path + '\\', fileName + '.xlsx']
    return fileAndPath


def insertEmailData(emailData,path):
    import openpyxl
    wb = openpyxl.load_workbook(path[0] + path[1])
    sheet = wb['Sheet']

    for i in range(len(emailData)):
        for j in range(len(emailData[i])):
            cell = chr(65 + j) + str(i + 2)
            if type(emailData[i][j]) == tuple:
                sheet[cell].value = emailData[i][j][0]
            else:
                sheet[cell].value = emailData[i][j]
                
    wb.template = False

    while 1:
        try:
            wb.save(path[1])
            print("\nThe excel file has been prepared at the chosen address.\nPress enter to exit the program.")
            input()
            break
        except PermissionError:
            print('The excel file is open. Please close the file and press enter to try again.')
            input()
            
print('''# ========================================== #
Welcome to the turnaround time analyzer!
Enter the name of your IMAP mail server.''')
plug = input()
print('\nNext, enter the email you want to analyze.')
address = input()
print('\nFinally, enter your password.')
pw = input()
print("\nLet's try that.")
print('# ========================================== #')

data = getAddAndSub(address,pw,plug)
path = createNewXlsx()
insertEmailData(data,path)


