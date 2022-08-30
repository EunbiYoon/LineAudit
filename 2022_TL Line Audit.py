import pandas as pd
import smtplib
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
import datetime

server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()

# 메일 내용 구성
msg = MIMEMultipart()

# 수신자 발신자 지정
msg['From'] = 'eunbi1.yoon@lge.com'
msg['To'] = 'russell.wilson@lge.com'
msg['Cc'] = 'iggeun.kwon@lge.com'
msg['Bcc'] = 'eunbi1.yoon@lge.com'

# Subject 꾸미기
today = date.today()
today = today.strftime('%m/%d')
msg['Subject'] = '[Daily report, ' + today + "] TN Factory Top Loader Line Audit by R&D"

today = date.today()
today = today.strftime('%m%d')

######################################################### Top Loader Line Audit Result #####################################################
data = pd.read_excel('//US-SO11-NA08765/R&D Secrets/GMES_Line Audit/3_TL Daily Line Audit/' + today + '.xlsx',
                     sheet_name=1)
# data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/GMES_Line Audit/3_TL Daily Line Audit/'+'0111.xlsx',sheet_name=1)
# data=data[data['Model'] == 'ACQ30063511']

data1 = data[['Insp. Items.1']]
data2 = data[['Insp. Items.3', 'Insp. Items.2']]
data3 = data[['Measure Value', 'Measure Value.1', 'Measure Value.2', 'Measure Value.3', 'Measure Value.4', 'Judgment']]

data1.columns = ["Items"]
data2.columns = ["LSL", "USL"]
data3.columns = ["Value 1", "Value 2", "Value 3", "Value 4", "Value 5", "Judge"]

data1 = data1.drop([0], axis=0)

data2 = data2.drop([0], axis=0)
data2 = data2.fillna("")

data3 = data3.drop([0], axis=0)
data3 = data3.fillna("OK")

data = pd.concat([data1, data2, data3], axis=1)

# 숫자만 추출 
data["No"] = data["Items"].str.extract(r"(\d+\.\d+|\d+)").astype('int')
data = data.set_index("No")
data = data.sort_index()

# 문자만 추출
data["Items"] = data["Items"].str.replace(r"(\d+\.\d+|\d+)", "")
data["Items"] = data["Items"].str.replace(pat=r'.', repl=r'', regex=True)

col_colours = ['#FBEEB0', '#C5EFE8', '#C5EFE8', '#CED6FE', '#CED6FE', '#CED6FE', '#CED6FE', '#CED6FE', '#FCD9E0']

###########################설정
# gs = fig.add_gridspec(24,10)

Today = datetime.datetime.now()
Firstday = datetime.datetime.strptime("20200101", "%Y%m%d")
Re_diff = (Today - Firstday).days

if (Re_diff % 2) == 0:
    print("Main Day")
    fig = plt.figure(constrained_layout=True, figsize=(10, 5))
    ax0 = fig.add_subplot()
    ax0.annotate('Daily Audit (Main Line) Result', xy=(0.095, 0.85), color='#515C5A', fontsize=10)
    text4 = "2. Daily Top Loader Line Audit Results \n   2-1) Main Line(14 check points)"

else:
    print("Sub Day")
    fig = plt.figure(constrained_layout=True, figsize=(10, 9))
    ax0 = fig.add_subplot()
    ax0.annotate('Daily Audit (Sub Line) Result', xy=(0.095, 0.85), color='#515C5A', fontsize=10)
    text4 = "2. Daily Top Loader Line Audit Results\n   2-1) Sub Line(22 check points, Image 1)"

ax0.set_axis_off()
data_table = ax0.table(cellText=data.values, rowLabels=data.index, colLabels=data.columns, loc='center',
                       colLoc='center', rowLoc='left', cellLoc='left', colColours=col_colours)
data_table.set_fontsize(9)
data_table.auto_set_column_width(col=list(range(len(data.columns))))

# plt.show()
plt.savefig('TL Line Audit.png')

###############################################################################################################################################3


# Body 첨부하기
text0 = 'This is DX activities from LGEUS R&D Team\nPerson in charge: LGEUS R&D Team Eunbi Yoon\n\n\n'
text1 = "Dear HA Quality Management Divison President,\n\n"
text2 = "I'd Like to report the Daily TN Factory Top Loader Line Audit by R&D"
text3 = "1. Issue Items : None"

textblank = '\n'

msg.attach(MIMEText(text0, 'plain'))
msg.attach(MIMEText(text1, 'plain'))
msg.attach(MIMEText(text2, 'plain'))
msg.attach(MIMEText(text3, 'plain'))
msg.attach(MIMEText(text4, 'plain'))

# 첨부 파일1
with open('TL Line Audit.png', 'rb') as f:
    img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('TL Line Audit.png'))
msg.attach(image)

msg.attach(MIMEText(textblank, 'plain'))

# 첨부 파일2
with open('sign.png', 'rb') as f:
    img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('sign.png'))
msg.attach(image)

# 메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")



