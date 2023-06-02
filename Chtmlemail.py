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

#Subject 꾸미기
msg['Subject']='[테네시 재료비 관리 Task] 5월 4주차 BOM과 실제 생산 투입 재료비 차이 분석'

# html table attach
FL_attach = MIMEText(FL_html, "html")
TL_attach = MIMEText(TL_html, "html")
DR_attach = MIMEText(DR_html, "html")
FL_attach_item = MIMEText(FL_item, "html")
TL_attach_item = MIMEText(TL_item, "html")
DR_attach_item = MIMEText(DR_item, "html")

msg.attach(MIMEText('<h4 style="font-weight:300;font-family:Arial Narrow; color:black">Dear All, <br/><br/>I would like to share TN Production Site 3 Main Model Material Cost Trend.<br/>Please refer to the attachment and below information.<br/>Thank you,<br/><br/></h4>','html'))

msg.attach(MIMEText('<h3 style="font-family:Arial Narrow; color:grey">Front Loader - F3P2CYUBW.ABWEUUS</h3>','html'))
msg.attach(FL_attach)
msg.attach(MIMEText('<h4 style="font-family:Arial Narrow; color:navy">- NPT vs GERP  Top 7 Difference Items','html'))
msg.attach(FL_attach_item)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:Arial Narrow; color:grey">Top Loader - T1889EFHUW.ABWEUUS</h3>','html'))
msg.attach(TL_attach)
msg.attach(MIMEText('<h4 style="font-family:Arial Narrow; color:navy">- NPT vs GERP  Top 7 Difference Items','html'))
msg.attach(TL_attach_item)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:Arial Narrow; color:grey">Dryer - RV13D1AMAZU.ABWEUUS</h3>','html'))
msg.attach(DR_attach)
msg.attach(MIMEText('<h4 style="font-family:Arial Narrow; color:navy">- NPT vs GERP  Top 7 Difference Items','html'))
msg.attach(DR_attach_item)


#첨부 파일1
etcFileName='FL_BOM_Comparison_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/BOM Comparison_FL.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

#첨부 파일2
etcFileName='TL_BOM_Comparison_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/BOM Comparison_TL.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

#첨부 파일3
etcFileName='DR_BOM_Comparison_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/BOM Comparison_DR.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)


#첨부 파일4
etcFileName='result_0526.xlsx'
with open('C:/Users/RnD Workstation/Documents/NPTGERP/0526/result_0526.xlsx', 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)


#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")