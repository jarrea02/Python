from win32com.client import Dispatch
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
root_folder = outlook.Folders.item(2)
#print root_folder
dev_data_req = root_folder.Folders['All Public Folders'].Folders['Administration'].Folders['Development Actions'].Folders['DevDataRequest']
#print dev_data_req

messages = dev_data_req.Items

mess = messages.Getlast
body_content = mess.body
mess_subject = mess.subject
#print body_content
#print mess.attachments.item().DisplayName




for i in messages:


    if i.attachments.count >= 1: #and i.attachments.item(1).DisplayName.find('png') == False:
       # f =  i.attachments.item(1).DisplayName
        #if f.find('image0') == -1:
          # print f
        for j in range(1,i.attachments.count):
            f = i.attachments.item(j).DisplayName
            if f.find('image0') == -1:
                print f
                print i.subject
                print i.ReceivedTime
                print i.SenderName
                3print i.Sender
                #print i.SenderEmailAddress
                print i.ConversationID
                print i.ConversationIndex
                #print i.TaskCompletedDate






