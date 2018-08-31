from win32com.client import constants
from win32com.client.gencache import EnsureDispatch as Dispatch

def list_folders( mapi_object, indent = '' ) :
  for _, folder in enumerate( mapi_object.Folders, 1 ) : # 1-index
    if folder.DefaultItemType == constants.olMailItem :
      print( '{}{}'.format( indent, folder.Name ) )
      list_folders( folder, indent = indent + ' ' )

#def to_folder(mapi, folder):


mapi = Dispatch("Outlook.Application").GetNamespace("MAPI")
#list_folders( mapi )
#mapi.Logon("Outlook")
#inbox = mapi.GetDefaultFolder()
inbox = mapi.Folders.Item("production")
print(inbox)
'''
messages = inbox.Items
message = messages.GetLast()
body = message.Body
print(body)
'''
