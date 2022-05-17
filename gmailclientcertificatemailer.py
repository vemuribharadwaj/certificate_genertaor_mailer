# Python code to illustrate Sending mail with attachments 
# from your Gmail account 

# libraries to be imported
import pandas as pd
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

#read data from excel into pandas
df = pd.read_excel('NewDataset.xlsx')
#print(df)
#print("Dimention : ",df.shape)
(rows, columns) = df.shape
#print("Number of rows = ",rows)
#print("Number of columns = ",columns)
fromaddr = "dummydoop85@gmail.com"#Place your mail ID Here
# creates SMTP session 
s = smtplib.SMTP('smtp.gmail.com', 587)#for gmail client 
# start TLS for security #find TLS 
s.starttls() 
# Authentication 
s.login("dummydoop85@gmail.com","12345@dD")#Enter your Password Here 

for i in range(rows):
    fileLocation = df.loc[i].CERTIFICATE_LOCAATION #Place the location of the file you want to send as an attachment
    toaddr = df.loc[i].EMAIL
    # instance of MIMEMultipart 
    msg = MIMEMultipart() 
    # storing the senders email address 
    msg['From'] = "TALENT MEET 2020<dummydoop85@gmail.com>"#In the angular braces place the sender mail address
    # storing the receivers email address 
    msg['To'] = toaddr 
    # storing the subject 
    msg['Subject'] = "Congratulations champ!!"#Type the subject of the mail Here
    # string to store the body of the mail 
    body = """Dear participant congratulations

    """#write the body of the mail you want to send
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 
    # open the file to be sent 
    filename = df.loc[i].NAME+".pdf"
    attachment = open(fileLocation, "rb") 
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
    # encode into base64 
    encoders.encode_base64(p) 
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 
    # Converts the Multipart msg into a string 
    text = msg.as_string() 
    # sending the mail 
    s.sendmail(fromaddr, toaddr, text)
    print(str(i+1)+". Mail sent to "+df.loc[i].NAME)
# terminating the session 
s.quit() 
