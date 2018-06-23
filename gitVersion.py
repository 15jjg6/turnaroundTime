#! python3

# The first function opens the input inbox,
# analyzes all of the messages in that inbox
# and outputs a list that includes:
# * reciever and sender addresses 
# * subject line
# * sent and recieve times

# Requires imapclient, imaplib, pyzmail, openpyxl, pprint, re

def getAddAndSub(username,pword,hostAddress):

    # Check these libraries to see which ones are redundant. 


    # Increases the max allowed size of program to 10,000,000
    # bytes, much less likely to get this error:
    # imaplib.error: got more than 10000 bytes
    import imaplib
    imaplib._MAXLINE = 10000000
    
    import re
    responseDateRegex = re.compile(r'Date: (\w\w\w, \d\d \w\w\w \d\d\d\d \d\d:\d\d:\d\d (-|\+)\d\d\d\d)')
    originalDateRegex = re.compile(r'Sent: ((Mon|Tues|Wednes|Thurs|Fri|Satur|Sun)day, (January|February|March|April|May|June|July|August|September|October|November|December) (\d\d|\d), \d\d\d\d (\d\d|\d):\d\d (A|P)M)')

    import imapclient
    with imapclient.IMAPClient(host=hostAddress) as client:

        # Login and access the inbox.
        try:
            client.login(username,pword)
        except imaplib.IMAP4.error:
            print('The password is incorrect. Please run the program again.')
            raise SystemExit
        client.select_folder('INBOX',readonly=True)

        # Aquire the message #'s of all messages in the inbox.
        messages = client.search(['ALL'])
        length = len(messages)

        print('Your password was correct. Looking through ' + str(length) + ' emails...')
        info = []
        import pyzmail
        for i in range(length):   
            # Consider cleaning this part up. Copying the entire
            # cache of emails is gonna take a lot of time. Is there
            # any way I can reduce the amount of info needed to be
            # copied? (Took about 45 seconds to run with 137 emails.)
            
            info.append(['','','','','',''])
            rawMessages = client.fetch(messages[i], ['BODY[]'])
            legibleMessage = pyzmail.PyzMessage.factory(rawMessages[messages[i]][b'BODY[]'])

            exitDate = responseDateRegex.search(str(legibleMessage))           
            recieveDate = originalDateRegex.search(str(rawMessages))
            
            info[i][0] = messages[i]
            info[i][1] = legibleMessage.get_address('from')
            info[i][2] = legibleMessage.get_address('to')
            info[i][3] = legibleMessage.get_subject()
            if recieveDate != None:
                info[i][4] = str(recieveDate.group(1))
            info[i][5] = exitDate.group(1)
            # Only transcribes the email of the first recipient.
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

    print('''\nYour new excel file with the turnaround data will be created next to this python script. What do you want to name this file?

Make sure this file doesn't have the same name as another excel spreadsheet in the same folder if you want to keep the old one, the script will save over it!''')
    fileName = input()
    print('Lets try that.')
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
    
    try:
        wb.save(path + '\\' + fileName + '.xlsx')
        print('The file ' + fileName + '.xslx has been saved in the above folder.')
    except PermissionError:
        print('There is another file in the same folder with the same name that is open.\nClose the file and try again if you are okay with saving over it.')
        createNewXlsx()
    fileAndPath = [path + '\\', fileName + '.xlsx']
    return fileAndPath


def insertEmailData(emailData,path):
    import openpyxl
    wb = openpyxl.load_workbook(path[0] + path[1])
    print(wb)
    sheet = wb['Sheet']

    print('\nExcel file opened.')
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
            print("The excel file has been prepared at the chosen address.")
            input()
            break
        except PermissionError:
            print('The excel file is open. Please close the file and press enter to try again.')
            repeat = input()
            
            
# Run this section as a test.
# ========================================== #
print('Welcome to the turnaround time analyzer! Please enter the password to begin.')
pw = input()
address = 'joe.grosso@cogeco.ca'
plug = 'imap.cogeco.ca'
data = getAddAndSub(address,pw,plug)
path = createNewXlsx()
insertEmailData(data,path)
# ========================================== #
