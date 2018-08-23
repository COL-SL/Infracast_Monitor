import win32com.client
import win32com
import os
import pickle
import time
from functions import *

outlook = win32com.client.Dispatch("outlook.Application").GetNameSpace("MAPI")
inbox = outlook.Folders("mobile.gpoc.businesssolutions@telefonica.com").Folders("GMSC").Folders("A2P").Folders("Infracast")
message = inbox.items
#message = message.GetLast()


LIST_NUMBER_GATEWAY = []
LIST_NAME_COUNTRY = []
LIST_COUNTRY_CODE = []

LIST_NUMBER_GATEWAY = get_number_gateway(message)
print (LIST_NUMBER_GATEWAY)

LIST_NAME_COUNTRY = get_name_country(message)
print (LIST_NAME_COUNTRY)

LIST_COUNTRY_CODE = get_country_code(message)
print (LIST_COUNTRY_CODE)

#while(1):

    #        print ('\n')
        #infolist.append((receipt, subject, sender))
# time.sleep(544)

'''
infolist = []
for message2 in message:
    #message2=message.GetLast()
    time.sleep(5)
    subject=message2.Subject
    #date1=message2.senton.Date()
    sender = message2.Sender
    receipt = message2.ReceivedTime
    body = message2.body
    #print (receipt, " | ", subject, " | ", sender, " | ", body)
    print (receipt)
    infolist.append((receipt, subject, sender))
    message2.Save
    message2.Close(0)
#fp = open("C:\Python27\\emails.pkl","w")
#pickle.dump(infolist, fp)
#fp.close()



from win32com.client import Dispatch
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
root_folder = outlook.Folders.Item(1)
print (root_folder.Name)

for folder in root_folder.Folders:
    print (folder.Name)
    print (objParentFolder.Folders(folder.Name))

inbox = root_folder.Folders['A2P']

print (inbox)
'''




'''
your_folder = mapi.Folders['GMSC'].Folders['A2P'].Folders['Infracast']
for message in your_folder.Items:
    print(message.Subject)
'''

'''

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
message = messages.GetLast()
body_content = message.body
print (body_content)

'''
