import win32com.client
import win32com
import os
import pickle
import time
from functions import *
import re

outlook = win32com.client.Dispatch("outlook.Application").GetNameSpace("MAPI")
inbox = outlook.Folders("mobile.gpoc.businesssolutions@telefonica.com").Folders("GMSC").Folders("A2P").Folders("Infracast")
message = inbox.items

list_number_gateway = []
list_name_country = []
list_country_code = []
list_normalize_country_code = []
list_high_rate = []
list_percent =[]
list_normalize_percent =[]
list_out =[]
list_normalize_out =[]
list_messages =[]
list_normalize_messages =[]
list_final = []
count_element_total = 0

list_number_gateway = get_number_gateway(message)
print(list_number_gateway)

list_country_code = get_country_code(message)
#print (LIST_COUNTRY_CODE)

list_normalize_country_code = get_normalize_country_code(list_country_code)
print(list_normalize_country_code)

list_name_country = get_name_country(list_normalize_country_code)
print(list_name_country)

list_high_rate = get_high_rate(message)
print(list_high_rate)

list_percent = get_percent(message)
#print(list_percent)

list_normalize_percent = get_normalize_percent(list_percent)
print(list_normalize_percent)

list_out = get_out(message)
#print(list_out)

list_normalize_out = get_normalize_out(list_out)
print(list_normalize_out)

list_messages = get_messages(message)
#print(list_messages)

list_normalize_messages = get_normalize_messages(list_messages)
print(list_normalize_messages)

count_element_total = count_element_list(list_normalize_messages )
print (count_element_total)

concat_list_final(count_element_total, list_number_gateway, list_normalize_country_code, list_name_country,
                               list_high_rate, list_normalize_percent, list_normalize_out, list_normalize_messages)



'''
pattern = re.compile('(\( *[0-9]+ *\))')
cadena = 'a44453'
patron.match(cadena)  # <_sre.SRE_Match object at 0x02303BF0>
patron.search(cadena) # <_sre.SRE_Match object at 0x02303C28>
cadena = 'ba3455' # la coincidencia no está al principio!
patron.search(cadena)  #  <_sre.SRE_Match object at 0x02303BF0>
#print (str(texto))
print (patron.findall(str(texto))) # None
'''
#buscar(['Gateway'], texto)
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
