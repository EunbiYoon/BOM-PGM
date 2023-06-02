import os
import pandas as pd
import numpy as np
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 
from Bhtmltable import FL_html,TL_html,DR_html,FL_item,TL_item,DR_item
from ADataDR import today_date

################################## Send email ################################## 
server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
# msg['To']='iggeun.kwon@lge.com, incheol.kang@lge.com, sehee.aiello@lge.com, jacey.jung@lge.com, gilnam.lee@lge.com, steven.yang@lge.com, jajoon1.koo@lge.com, wolyong.ha@lge.com, dowan.han@lge.com'
# msg['Cc']='ethan.son@lge.com, jongseop.kim@lge.com, richard.song@lge.com, minhyoung.sun@lge.com, kitae3.park@lge.com, tg.kim@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'

msg.attach(MIMEText('<h4 style="font-family:sans-serif; font-weight:500">Dear all,<br/><br/>I would like to share this week cost review report and detailed informations are in below website.<br/>You can access website with Chrome or Edge in CloudPC or LG wifi for security purpose. <a href="http://10.225.2.85">http://10.225.2.85</a></h4>','html'))

save_path='C:/Users/RnD Workstation/Documents/CostReview/'+today_date+'/'

# graph file read
with open(save_path+'FL.JPG', 'rb') as f:
        img_data = f.read()
image1 = MIMEImage(img_data, name=os.path.basename("FL"+today_date+'.jpg'))

with open(save_path+'TL.JPG', 'rb') as f:
        img_data = f.read()
image2 = MIMEImage(img_data, name=os.path.basename("TL"+today_date+'.jpg'))

with open(save_path+'DR.JPG', 'rb') as f:
        img_data = f.read()
image3 = MIMEImage(img_data, name=os.path.basename("DR_BPA_Entity"+today_date+'.jpg'))



# msg.attach(MIMEText('<br/><h3 style="font-family:sans-serif;">Front Loader BPA Entity Trend</h3>','html'))
msg.attach(image1)
msg.attach(image2)
msg.attach(image3)



#첨부 파일1
etcFileName='Cost Review_0602.xlsx'
with open("C:/Users/RnD Workstation/Documents/CostReview/0602/Cost Review_0602.xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)



#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")
