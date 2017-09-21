import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate
from email import encoders
import mimetypes

def send_mail_file(send_from, send_to, send_cc, subject, filename):
    server="appmail.risk.regn.net"
    fp = open(filename, 'r')
    print ("Are we here ?????\n\n\n")
    #print (fp.read())
    msg = fp.read()
    print(type(msg))
    fp.close()
    assert type(send_to)==list
    
    print("What is the MSG : {0}".format(msg))
    
    MESSAGE = MIMEMultipart('alternative')
    
    MESSAGE['subject'] = subject
    
    MESSAGE['To'] = COMMASPACE.join(send_to)
    MESSAGE['From'] = send_from
    
    print (send_to)
    print(send_from)
    
    HTML_BODY = MIMEText(msg, 'html')
    MESSAGE.attach(HTML_BODY)
    
    smtp = smtplib.SMTP(server)
    smtp.sendmail(send_from, send_to, MESSAGE.as_string())
    smtp.close()
    

def send_mail(send_from, send_to, send_cc, subject, text):
    server="appmail.risk.regn.net"
    if send_cc == '':
       send_cc = []
    print('send_from')
    print(send_from)
    print('send_to')
    print(send_to)
    print(send_to)
    assert type(send_to)==list
    assert type(send_cc)==list

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Cc'] = COMMASPACE.join(send_cc)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach( MIMEText(text) )

    # for f in files:
        # print('f is .. ', f)
        # part = MIMEBase('application', "octet-stream")
        # part.set_payload( open(f,"rb").read() )
        # Encoders.encode_base64(part)
        # part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        # msg.attach(part)
        # print(os.path.basename(f))

    smtp = smtplib.SMTP(server)
    smtp.sendmail(send_from, send_to + send_cc, msg.as_string())
    smtp.close()
