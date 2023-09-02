import requests
import datetime
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# 現在時刻を取得する
dt_now = datetime.datetime.now()

# 気象庁のアジア天気図を取得して、保存する
def jma_get_img(get_date):
    data_fmt = 'pdf'
    jma_file_name = 'jma/jma-' + str(get_date) + '.' + data_fmt

    jma_url = 'https://www.data.jma.go.jp/yoho/data/wxchart/quick/' + get_date[:6] + '/ASAS_COLOR_' + get_date + '.' + data_fmt
    jma_data = requests.get(jma_url).content

    with open(jma_file_name, 'wb') as file:
        file.write(jma_data)

# 波浪情報
def imoc_get_img():
    imoc_url = 'https://www.imocwx.com/cwm/cwmjp_00.png'
    imoc_data = requests.get(imoc_url).content
    imoc_file_name = 'imoc/imoc-' + dt_now.strftime('%Y%m%d') + '.png'

    with open(imoc_file_name, 'wb') as file1:
        file1.write(imoc_data)

# メールサーバに送信を実行させる
def send_outlook_mail(msg):
    """
    引数msgをOutlookで送信
    """
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    # ログインしてメール送信
    server.login(my_account, my_password)
    server.send_message(msg)

# メールの電文を作成する
def make_mime(mail_to, subject, body):
    """
    引数をMIME形式に変換
    """
    msg = MIMEMultipart()
    # msg = MIMEText(body, 'plain') #メッセージ本文
    msg['Subject'] = subject #件名
    msg['To'] = mail_to #宛先
    msg['From'] = my_account #送信元
    msg.attach(MIMEText(body, 'plain'))

    # 天気図の添付
    with open("jma/jma-202309020000.pdf", "rb") as f:
        jma_attachment = MIMEApplication(f.read())
    jma_attachment.add_header("Content-Disposition", "attachment", filename="jma-202309020000.pdf")
    msg.attach(jma_attachment)

    # 波浪図の添付
    with open("imoc/imoc-20230902.png", "rb") as f:
        imoc_attachment = MIMEApplication(f.read())
    imoc_attachment.add_header("Content-Disposition", "attachment", filename="imoc-20230902.png")
    msg.attach(imoc_attachment)

    return msg

# メールを送信する
def send_my_message():
    """
    メイン処理
    """
    # MIME形式に変換
    msg = make_mime(
        mail_to='yoshizumi@zoho.com', #送信したい宛先を指定
        subject='テスト件名',
        body='テストです。テストです。テストです。')
    # gmailに送信
    send_outlook_mail(msg)


# メイン関数
if __name__ == '__main__':

    # Outlook設定
    my_account = 'xxxx@xxxxx.xx'
    my_password = 'password'

    # 気象庁の天気図を取得して、保存する
    get_date = dt_now.strftime('%Y%m%d0000')
    jma_get_img(get_date=get_date)

    # 波浪図を取得する
    imoc_get_img()

    # メールを送信する
    send_my_message()
