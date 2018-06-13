# This function will open the input inbox,
# analyze all of the messages in that inbox
# and output a tuple or dictionary with the
# addresses of the reciever and the sender,
# and it will have the subject of the email.

# To improve this function in the future, I
# will add the functionality to find the times
# of the emails. 

# Requires imapclient, imaplib, pyzmail, openpyxl, pprint.

def getAddAndSub(username,pword,hostAddress):

    # Check these libraries to see which ones are redundant. 
    import imapclient, imaplib, pyzmail

    # Increases the max allowed size of program to 10,000,000
    # bytes, much less likely to get this error:
    # imaplib.error: got more than 10000 bytes
    # imaplib._MAXLINE = 10000000

    with imapclient.IMAPClient(host=hostAddress) as client:

        # Login and access the inbox.
        client.login(username,pword)
        client.select_folder('INBOX',readonly=True)

        # Aquire the message #'s of all messages in the inbox.
        messages = client.search(['ALL'])
        
        info = []
        for i in range(len(messages)):
            print(i)
            # Consider cleaning this part up. Copying the entire
            # cache of emails is gonna take a lot of time. Is there
            # any way I can reduce the amount of info needed to be
            # copied? (Took about 45 seconds to run with 137 emails.)
            info.append(['','','',''])
            rawMessages = client.fetch(messages[i], ['BODY[]'])
            legibleMessage = pyzmail.PyzMessage.factory(rawMessages[messages[i]][b'BODY[]'])

            info[i][0] = messages[i]
            info[i][1] = legibleMessage.get_address('from')
            info[i][2] = legibleMessage.get_address('to')
            info[i][3] = legibleMessage.get_subject()
            # Only transcribes the email of the first recipient. 
        client.logout()
        
    return info
    # Output is the unique IDs (UIDs) subject, sender
    # and reciever of each email.

def insertEmailData(emailData):
    import openpyxl
    wb = openpyxl.load_workbook("C:\\Users\\example\\Documents\\quoteTurnaroundTimes.xlsx")
    sheet = wb['Sheet1']
    for i in range(len(emailData)):
        for j in range(len(emailData[i])):
            cell = chr(65 + j) + str(i + 2)
            if type(emailData[i][j]) == tuple:
                sheet[cell].value = emailData[i][j][0]
            else:
                sheet[cell].value = emailData[i][j]
    wb.template = False
    wb.save(r"quoteTurnaroundTimes.xlsx")
    print("The excel file has been prepared at the chosen address.")



# Run this section as a test.
# ========================================== #
import pprint
address = 'example@gmail.com'
password = '123456'
plug = 'imap.gmail.com'
data = getAddAndSub(address,password,plug)
pprint.pprint(data)
insertEmailData(data)
# ========================================== #
