# Test
import os
import base64

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, Attachment, FileContent, FileName, FileType, Disposition)

##API_KEY = "SG._2Xep5VXQSOt8PecYR6OYw.7a9B92QKKSB10vP56Hek9eQKPvnWmYwn22nMeRF0AQ0SG._2Xep5VXQSOt8PecYR6OYw.7a9B92QKKSB10vP56Hek9eQKPvnWmYwn22nMeRF0AQ0"
API_KEY = "SG.xeEYdb_SQCOkkemn_KKarQ.3k6wPY8rOJsWXwklMTO28lO_IZHVH-Gqh3lUPlyn344SG.xeEYdb_SQCOkkemn_KKarQ.3k6wPY8rOJsWXwklMTO28lO_IZHVH-Gqh3lUPlyn344"
message = Mail(
    from_email='superadmin@vflowtech.com',
    to_emails=', '.join(['superadmin@vflowtech.com']),
    subject='Test Status Report',
    html_content='<strong>Pioneering Tomorrow</strong>'
)

with open('test.docx', 'rb') as f:
    data = f.read()
    f.close()
encoded_file = base64.b64encode(data).decode()

attachedFile = Attachment(
    FileContent(encoded_file),
    FileName('test'),
    FileType('application/docx'),
    Disposition('attachment')
)
message.attachment = attachedFile

sg = SendGridAPIClient(API_KEY)
response = sg.send(message)
print(response.status_code, response.body, response.headers)




