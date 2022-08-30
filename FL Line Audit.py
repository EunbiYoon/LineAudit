import pandas as pd
import pandas as pd
import numpy as np
import xlrd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import msoffcrypto
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
import re
import datetime

server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To']='russell.wilson@lge.com'
msg['Cc']='iggeun.kwon@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'
#msg['To']='sujong.lee@lge.com'
#msg['Cc']='steve.baek@lge.com, youngsooo.kim@lge.com, daewook.kwak@lge.com, jichang.son@lge.com, sangho.an@lge.com, sunggi.hwang@lge.com, antony.jung@lge.com, youngbae.park@lge.com, jungyu.kim@lge.com, donggil.lee@lge.com, deuko.kim@lge.com, jiyoon1.heo@lge.com, taehodream.lee@lge.com, cash.yun@lge.com, sanghwa.kim@lge.com, jihoon81.kim@lge.com, jaeeun.chung@lge.com, dan.roach@lge.com, si1207.kim@lge.com, hyungjin.jung@lge.com, charles.lonergan@lge.com, jounghun.han@lge.com, bronson.allen@lge.com, minwoo2122.lee@lge.com, christine.broadhurst@lge.com, corey.baynham@lge.com, patrick.stevenson@lge.com, chris.cole@lge.com, david8.kim@lge.com, iggeun.kwon@lge.com, aaron1.garcia@lge.com, soonan.park@lge.com, eunbi1.yoon@lge.com, dharmin.mistry@lge.com, min1.park@lge.com, remoun.abdo@lge.com, russell.wilson@lge.com'

#Subject 꾸미기
today=date.today()
today=today.strftime('%m/%d')
msg['Subject']='[Daily report, '+today+"] TN Factory Front Loader Line Audit by R&D"

today=date.today()
today=today.strftime('%m%d')
print(today)

############## plot의 요소들을 하나로 묶기
fig, axes = plt.subplots(2,2)
fig.set_size_inches(8,6)

#################################################################### Balance Weight #########################################################3
############## 1x1 chart
# 데이터 불러오기
bw_data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/GMES_Line Audit/4_FL Daily Line Audit/'+today+'.xlsx',sheet_name=2)
print(bw_data)

# Box Plot
BW_Box=axes[0,0].boxplot(bw_data)
axes[0,0].set_title('Balance Weight Bolt Torque',fontsize=10)
today_weekday=datetime.datetime.today().weekday()

# reference line 변수 설정
hline_70=1.5+today_weekday
hline_120=1.5+today_weekday
annotate_70=1.5+today_weekday
annotate_120=1.5+today_weekday

axes[0,0].hlines(xmin=0, xmax=hline_70, y=70,color='r', linestyles='--')
axes[0,0].hlines(xmin=0, xmax=hline_120, y=120,color='r', linestyles='--')
axes[0,0].annotate('70', xy=(annotate_70, 69),ha='right', va='top',color='red',fontsize=9)
axes[0,0].annotate('120', xy=(annotate_120, 119),ha='right', va='top',color='red',fontsize=9)


bwcolumns=bw_data.columns.strftime('%m/%d')
col_labels=bwcolumns

axes[0,0].set_xticklabels(col_labels, fontsize=9)

axes[0,0].set_ylim(60,130,10)
axes[0,0].set_xlabel('Date',fontsize=9,color='gray')
axes[0,0].set_ylabel('kgf*cm',fontsize=9,color='gray')

############# 2x1 Table
bw_max=round(bw_data.max(),2)
bw_min=round(bw_data.min(),2)
bw_mean=round(bw_data.sum()/30,2)
bw_table=pd.DataFrame([bw_max,bw_min,bw_mean])
bw_table.index=["Maximum","Minimum","Mean"]

axes[1,0].set_axis_off()
cell_text=bw_table.values
row_labels=["Maximum","Minimum","Average"]


row_colours=["#B9EBD6","#B9EBD6","#B9EBD6"]
col_colours=np.full(today_weekday+1,"#B9EBD6")

BW_TABLE=axes[1,0].table(cellText=bw_table.values, rowLabels=row_labels, colLabels=col_labels, loc='center', rowColours=row_colours, colColours=col_colours, cellLoc='center')
BW_TABLE.auto_set_font_size(False)
BW_TABLE.set_fontsize(9)
BW_TABLE.auto_set_column_width(col=list(range(len(bw_table.columns))))
axes[1,0].set_title('Balance Weight Bolt Torque Result',x=0.35,y=0.7,fontsize=10)

#################################################################### Door Latch #########################################################3
############## 1x2 chart
# 데이터 불러오기
dl_data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/GMES_Line Audit/4_FL Daily Line Audit/'+today+'.xlsx',sheet_name=3)

# Box Plot
axes[0,1].boxplot(dl_data)
axes[0,1].set_title('Door Latch Gap',fontsize=10)

# reference line 변수 설정
hline_70=1.5+today_weekday
hline_120=1.5+today_weekday
annotate_70=1.5+today_weekday
annotate_120=1.5+today_weekday


axes[0,1].hlines(xmin=0, xmax=hline_70, y=3.8,color='r', linestyles='--')
axes[0,1].hlines(xmin=0, xmax=hline_120, y=5.8,color='r', linestyles='--')
axes[0,1].annotate('3.8', xy=(annotate_70, 3.78),ha='right', va='top',color='red',fontsize=9)
axes[0,1].annotate('5.8', xy=(annotate_120, 5.78),ha='right', va='top',color='red',fontsize=9)


dlcolumns=dl_data.columns.strftime('%m/%d')
col_labels=dlcolumns

axes[0,1].set_xticklabels(col_labels,fontsize=9)
axes[0,1].set_ylim(3.5,6,0.5)
axes[0,1].set_xlabel('Date',fontsize=9,color='gray')
axes[0,1].set_ylabel('mm',fontsize=9,color='gray')

############# 2x2 Table
dl_max=round(dl_data.max(),2)
dl_min=round(dl_data.min(),2)
dl_mean=round(dl_data.sum()/5,2)
dl_table=pd.DataFrame([dl_max,dl_min,dl_mean])
dl_table.index=["Maximum","Minimum","Average"]

axes[1,1].set_axis_off()
cell_text=dl_table.values
row_labels=["Maximum","Minimum","Average"]
row_colours=["#B9EBD6","#B9EBD6","#B9EBD6"]
col_colours=np.full(today_weekday+1,"#B9EBD6")
DL_TABLE=axes[1,1].table(cellText=dl_table.values, rowLabels=row_labels, colLabels=col_labels ,loc='center', rowColours=row_colours, colColours=col_colours, cellLoc='center')
DL_TABLE.auto_set_font_size(False)
DL_TABLE.set_fontsize(9)
DL_TABLE.auto_set_column_width(col=list(range(len(dl_table.columns))))
axes[1,1].set_title('Door Latch Gap Result',x=0.35,y=0.7,fontsize=10)


#그래프 간격 띄우기
plt.tight_layout()

# save fig
plt.savefig('FL Line Audit1.png')

######################################################### GMES table 결과 도출 #####################################################
#data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/GMES_Line Audit/4_FL Daily Line Audit/'+today+'.xlsx',sheet_name=1)# 수정
data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/GMES_Line Audit/4_FL Daily Line Audit/'+today+'.xlsx',sheet_name=1)
#data=data[data['Model'] == 'ACQ30063511']

data1=data[['Insp. Items.1']]
data2=data[[ 'Insp. Items.3','Insp. Items.2']]
data3=data[['Measure Value','Measure Value.1','Measure Value.2','Measure Value.3','Measure Value.4','Judgment']]

data1.columns=["Items"]
data2.columns=["LSL","USL"]
data3.columns=["Value 1","Value 2","Value 3","Value 4","Value 5", "Judge"]

data1=data1.drop([0],axis=0)

data2=data2.drop([0],axis=0)
data2=data2.fillna("")

data3=data3.drop([0],axis=0)
data3=data3.fillna("OK")

data=pd.concat([data1, data2,data3], axis=1)


#숫자만 추출 
data["No"]=data["Items"].str.extract(r"(\d+\.\d+|\d+)").astype('int')
data=data.set_index("No")
data=data.sort_index()

#문자만 추출
data["Items"]=data["Items"].str.replace(r"(\d+\.\d+|\d+)","")
data["Items"]=data["Items"].str.replace(pat=r'.', repl=r'', regex=True) 

col_colours=['#FBEEB0','#C5EFE8','#C5EFE8','#CED6FE','#CED6FE','#CED6FE','#CED6FE','#CED6FE','#FCD9E0']
    
fig, ax = plt.subplots(1,1)
fig.set_size_inches(8,4)
#fig.set_size_inches(8,7)
ax.set_axis_off()

data_table=ax.table(cellText=data.values,rowLabels=data.index,colLabels=data.columns, loc='center',colLoc='center',rowLoc='left',cellLoc='left',colColours=col_colours)
data_table.set_fontsize(9)
data_table.auto_set_column_width(col=list(range(len(data.columns))))

today=date.today()
today=today.strftime('%m/%d')
ax.set_title('Front Loader Line Audit Result ('+today+")",x=0.5, y=1.05,fontsize=10)


#그래프 간격 띄우기
plt.tight_layout()

# save fig
plt.savefig('FL Line Audit2.png')

###############################################################################################################################################3


#Body 첨부하기
text0='This is DX activities from LGEUS R&D Team\nPerson in charge: LGEUS R&D Team Eunbi Yoon\n\n'
text1="Dear HA Quality Management Divison President,\n"
text2="I'd Like to report the Daily TN Factory Front Loader Line Audit by R&D"
text3="1. Issue Items : None"
text4="2. Balance Weight Bolt Torque and Door Latch Gap Measurements (Image 1)"
text5="3. Daily Front Loader Line Audit Results (Image 2)\n\n\n"
text6="[Image 1] Boxplot on Balance Weight Bolt Torque and Door Latch Gap"
msg.attach(MIMEText(text0,'plain'))
msg.attach(MIMEText(text1,'plain'))
msg.attach(MIMEText(text2,'plain'))
msg.attach(MIMEText(text3,'plain'))
msg.attach(MIMEText(text4,'plain'))
msg.attach(MIMEText(text5,'plain'))
msg.attach(MIMEText(text6,'plain'))

#첨부 파일1
with open('FL Line Audit1.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('FL Line Audit1.png'))
msg.attach(image)

text7="[Image 2] Daily Audit Result (Cycle through two checklists. The combine checklists, include all 26 checkpoints.)"
msg.attach(MIMEText(text7,'plain'))

#첨부 파일1
with open('FL Line Audit2.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('FL Line Audit2.png'))
msg.attach(image)


#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")



