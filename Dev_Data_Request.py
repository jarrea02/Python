from win32com.client import Dispatch
from datetime import datetime

#from datetime import time
from datetime import date
import smartsheet
import logging
import os
import win32api
import string
import re
import time
logging.basicConfig()

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

root_folder = outlook.Folders['Public Folders - Adrian.Jarrett@mountsinai.org']

dev_data_req = root_folder.Folders['All Public Folders'].Folders['Administration'].Folders['Development Actions'].Folders['DevDataRequest']


ss_client = smartsheet.Smartsheet('ad16uh0f8n82it1h6mfr8nly38')
messages = dev_data_req.Items

request_sheet_ID = 1329883608573828

con_ID = 5646566888368004
sub_ID = 3394767074682756
sender_ID = 7898366702053252
received_ID = 580017307576196
con_index_ID = 5083616934946692
ss_created_ID = 887459656558468
ss_link_ID = 5391059283928964

sheet = ss_client.Sheets.get_sheet(
    request_sheet_ID)



path = 'K:/Raisers Edge QC Team/DevDataRequest/'
for i in messages:
        has_attachment = 'N'

        for j in range(1,i.attachments.count):
                FileName = i.attachments.item(j).DisplayName
                if FileName.find('image0') == -1:
                    has_attachment = 'Y'

        if has_attachment == 'Y':
        # file_path = 'K:\\Raisers Edge QC Team\\DevDataRequest'

            #print str(i.ReceivedTime)[0:8]


            regex = re.compile('([^\s\w]|_)+')
            n = regex.sub('', i.subject)
            t = regex.sub('', str(i.ReceivedTime)[0:8])
            r = regex.sub('', i.SenderName)
            i.SaveAs(path + t + ' ' + n + ', ' + r + '.msg')

            row_a = ss_client.models.Row()
            row_a.to_top = True
            row_a.cells.append({
                    'column_id': con_ID,
                    'value': i.ConversationID
                    })
            row_a.cells.append({
                    'column_id': sub_ID,
                    'value': i.subject
                })
            row_a.cells.append({
                    'column_id': sender_ID,
                    'value': i.SenderName
                    })
            row_a.cells.append({
                    'column_id': received_ID,
                'value': str(i.ReceivedTime)[0:8]
                })
      #  row_a.cells.append({
      #          'column_id': con_index_ID,
       #         'value': i.ConversationIndex
        #       })
            response = ss_client.Sheets.add_rows(
                request_sheet_ID,       # sheet_id
                [row_a])

            row_id = response.to_dict().get('data')[0].get('id')

            updated_attachment = ss_client.Attachments.attach_file_to_row(
                request_sheet_ID,
                row_id,
                (t + ' ' + n + ', ' + r + '.msg',
                open(path + t + ' ' + n + ', ' + r + '.msg', 'rb'),
                'application/msoutlook'
                )


            )

            time.sleep(6)
