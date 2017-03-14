import os
import csv 
import smtplib

cred_detail = []
os.chdir("c:\\data")
for row in csv.reader(open("pass.txt","rb")):       
        cred_detail.append(row)
username = cred_detail[0][0]
password = cred_detail[1][0]

fromaddr = username
toaddrs  = 'some.body@gmail.com'
msg = "\r\n".join([
  "From: Admin User <" + username + ">",
  "To: " + toaddrs,
  "Subject: Just a message",
  "",
  "Why, oh why"
  ])
server = smtplib.SMTP('smtp.gmail.com:587')
server.ehlo()
server.starttls()
server.login(username,password)
server.sendmail(fromaddr, toaddrs, msg)
server.quit()
