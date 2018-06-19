# This function will open the input inbox,
# analyze all of the messages in that inbox
# and output a tuple or dictionary with the
# addresses of the reciever and the sender,
# and it will have the subject of the email.

# Requires imapclient, imaplib, pyzmail, openpyxl, pprint, re

def getAddAndSub(username,pword,hostAddress):

    # Check these libraries to see which ones are redundant. 
    import imapclient, imaplib, pyzmail, re

    # Increases the max allowed size of program to 10,000,000
    # bytes, much less likely to get this error:
    # imaplib.error: got more than 10000 bytes
    imaplib._MAXLINE = 10000000
    
    responseDateRegex = re.compile(r'Date: (\w\w\w, \d\d \w\w\w \d\d\d\d \d\d:\d\d:\d\d (-|\+)\d\d\d\d)')

    originalDateRegex = re.compile(r'Sent: ((Mon|Tues|Wednes|Thurs|Fri|Sat|Sun)day, (January|February|March|April|May|June|July|August|September|October|November|December) (\d\d|\d), \d\d\d\d (\d\d|\d):\d\d (A|P)M)\n')
    
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
        
        print('Looking through ' + str(length) + ' emails...')
        info = []
        for i in range(length):   
            # Consider cleaning this part up. Copying the entire
            # cache of emails is gonna take a lot of time. Is there
            # any way I can reduce the amount of info needed to be
            # copied? (Took about 45 seconds to run with 137 emails.)
            
            info.append(['','','','','',''])
            rawMessages = client.fetch(messages[i], ['BODY[]'])
            legibleMessage = pyzmail.PyzMessage.factory(rawMessages[messages[i]][b'BODY[]'])

            exitDate = responseDateRegex.search(str(legibleMessage))            
            recieveDate = originalDateRegex.search(str(legibleMessage))
     
            info[i][0] = messages[i]
            info[i][1] = legibleMessage.get_address('from')
            info[i][2] = legibleMessage.get_address('to')
            info[i][3] = legibleMessage.get_subject()

            if recieveDate != None:
                info[i][4] = str(recieveDate.group(1))
            info[i][5] = exitDate.group(1)
            # Only transcribes the email of the first recipient.
            if (i+1)%10 == 0:
                print(str(i + 1) + '/' + str(length) + ' messages analyzed.')
        client.logout()
        print('All messages analyzed.')
    return info
    # Output is the unique IDs (UIDs) subject, sender
    # and reciever of each email, and times that the
    # quotes were recieved and responded to.

def insertEmailData(emailData):
    import openpyxl
    wb = openpyxl.load_workbook("C:\\Users\\path\\turnaroundTimes.xlsx")
    sheet = wb['Sheet1']
    print('\nExcel file opened.')
    for i in range(len(emailData)):
        for j in range(len(emailData[i])):
            cell = chr(65 + j) + str(i + 2)
            if type(emailData[i][j]) == tuple:
                sheet[cell].value = emailData[i][j][0]
            else:
                sheet[cell].value = emailData[i][j]
                
    wb.template = False
    try:
        wb.save(r"turnaroundTimes.xlsx")
    except PermissionError:
        print('The excel file is open. Please close the file and try again.')
        raise SystemExit
    print("The excel file has been prepared at the chosen address.")



# Run this section as a test.
# ========================================== #
import pprint
print('Welcome to the turnaround time analyzer! Please enter the password to begin.')
pw = input()
address = 'example123@gmail.com'
plug = 'imap.gmail.ca'
data = getAddAndSub(address,pw,plug)
# pprint.pprint(data)
insertEmailData(data)
# ========================================== #
