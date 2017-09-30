# Usage:
# Run the script from a folder with file "all.mbox"
# Attachments will be extracted into subfolder "attachments" 
# with prefix "n " where "n" is an order of attachment in mbox file. 
 
import mailbox, pickle, traceback, os
from email.header import decode_header
 
mb = mailbox.mbox('all.mbox')
 
prefs_path = '.save-attachments'
save_to = 'attachments/'

months = {}
months["Jan"] = "01"
months["Feb"] = "02"
months["Mar"] = "03"
months["Apr"] = "04"
months["May"] = "05"
months["Jun"] = "06"
months["Jul"] = "07"
months["Aug"] = "08"
months["Sep"] = "09"
months["Oct"] = "10"
months["Nov"] = "11"
months["Dec"] = "12"
 
if not os.path.exists(save_to): os.makedirs(save_to)
 
prefs = dict(start=0)
 
total = 0
failed = 0
 
def save_attachments(mid):
    msg = mb.get_message(mid)
    if msg.is_multipart():
        DateReceived = msg['Date'].split()[3]+'-'+months[msg['Date'].split()[2]]
        if int(msg['Date'].split()[1])<10 and len(msg['Date'].split()[1]) == 1:
            DateReceived = DateReceived+'-0'+msg['Date'].split()[1]
        else:
            DateReceived = DateReceived+'-'+msg['Date'].split()[1]
        for part in msg.get_payload():
            if str(part.get_filename()) == 'None':
                continue
             
            global total
            total = total + 1
 
            print()
            try:
                decoded_name = decode_header(part.get_filename())
                print(decoded_name)
                 
                if isinstance(decoded_name[0][0], str):
                    name = decoded_name[0][0]
                else:
                    name_encoding = decoded_name[0][1]
                    name = decoded_name[0][0].decode(name_encoding)
                 
                name = '%s %s %s' % (DateReceived, total, name)
                print('Saving %s' % (name))
                with open(save_to + name, 'wb') as f:
                    f.write(part.get_payload(decode=True))
            except:
                traceback.print_exc()
                global failed
                failed = failed + 1
 
 
for i in range(prefs['start'], 1000000):
    try:
        save_attachments(i)
    except KeyError:
        break
prefs['start'] = i
 
print()
print('Total:  %s' % (total))
print('Failed: %s' % (failed))