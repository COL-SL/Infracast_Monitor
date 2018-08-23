import win32com.client


from win32com.client import Dispatch
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
root_folder = outlook.Folders.Item(1)
print (root_folder.Name)

for folder in root_folder.Folders:
    print (folder.Name)

inbox = root_folder.Folders['FYI']

print (inbox)





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

