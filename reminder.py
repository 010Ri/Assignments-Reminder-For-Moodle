# seleniumを利用して動的Webページをスクレイピング
# 取得した提出課題情報から自動でoutlookメールを作成して送信

# import
import serial
import schedule
import time
import datetime
import win32com.client
from selenium import webdriver
from time import sleep
from selenium.webdriver.chrome.options import Options

# ログインに必要な要素
USERNAME = 'USERNAME'
PASSWORD = 'PASSWORD'

# メールに必要な情報：宛先
MailAccount = 'MAIL_ADRESS'

# get_information関数の定義：日にちを引数として、その日の提出課題の情報を入手し、Arduinoへ情報を送る
def get_information():
    # 初期設定
    target_url = 'Moodle URL'
    ser = serial.Serial('COM5',9600)
    error_flg = False
    assign = 0
    sub = 0
    count = 0
    tmp = 0

    # 現在の日付
    DateTime = datetime.date.today()
    year = str(DateTime.year)
    month = str(DateTime.month)
    day = str(DateTime.day)
    date = (year + '年 ' + month + '月 ' + day + '日 ')

    # WebDriverの起動
    options = Options()
    #options.add_argument('--headless')  # ここをコメントアウトすればヘッドレスモードを無効にできる
    driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\chromedriver',options=options)
    driver.get(target_url)
    sleep(3)

    # ログイン処理
    if error_flg is False:
        try:
            username_input = driver.find_element_by_name("username")
            username_input.send_keys(USERNAME)
            sleep(1)
 
            password_input = driver.find_element_by_name("password")
            password_input.send_keys(PASSWORD)
            sleep(1)
 
            username_input.submit()
            sleep(1)
        
        except Exception:
            print('ユーザー名、パスワード入力時にエラーが発生しました。')
            error_flg = True

    if mode == '1':
        year = str(DateTime.year)
        month = str(DateTime.month)
        day = str(DateTime.day)
        year = year + '年 '
        if month < '10':
            month = '0' + month + '月 '
        else:
            month = month + '月 '
        if day < '10':
            day = '0' + day + '日'
        else:
            day = day + '日'
        date = year + month + day
        print('検索する日付： ' + date)

        # 提出課題の情報を入手
        if error_flg is False:
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # javascriptを実行してページの最下部へ移動
                sleep(3)
                access_button = driver.find_element_by_xpath("//td[contains(@class, 'day text-center')]/a[contains(@data-title, '" + date + "')]").click()
                sleep(3)
            except Exception:
                error_flg = True
                print(date + 'の予定はありませんでした。')

        for elem_img in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/img[@src='path']"):
            if count == 0:
                print(date + 'の予定を取得しています。')
                for elem_h3 in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/h3"):
                    assign_list.append(elem_h3.text)
                for elem_span in driver.find_elements_by_xpath("//div/span[@class='date pull-xs-right m-r-1']"):
                        deadline_list.append(elem_span.text)
                for elem_a in driver.find_elements_by_xpath("//div[contains(@class, 'description card-block calendar_event')]/a"):
                    url_list.append(elem_a.get_attribute('href'))
            count += 1
        assign = count

        if count == 0:
            for elem_img in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/img[@src='path']"):
                if count == 0:
                    print(date + 'の予定を取得しています。')
                    for elem_h3 in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/h3"):
                        assign_list.append(elem_h3.text)
                    for elem_span in driver.find_elements_by_xpath("//div/span[@class='date pull-xs-right m-r-1']"):
                            deadline_list.append(elem_span.text)
                    for elem_a in driver.find_elements_by_xpath("//div[contains(@class, 'description card-block calendar_event')]/a"):
                        url_list.append(elem_a.get_attribute('href'))
                count += 1

        if error_flg is True:
            assign_list.append('\n')
            deadline_list.append(date + 'の予定はありません。')
            url_list.append('\n')

    if mode == '2':
        for ct in range(3):
            year = str(DateTime.year)
            month = str(DateTime.month)
            day = str(DateTime.day)
            year = year + '年 '
            if month < '10':
                month = '0' + month + '月 '
            else:
                month = month + '月 '
            if day < '10':
                day = '0' + day + '日'
            else:
                day = day + '日'
            date = year + month + day
            print('検索する日付： ' + date)

            # 提出課題の情報を入手
            if error_flg is False:
                try:
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # javascriptを実行してページの最下部へ移動
                    sleep(3)
                    access_button = driver.find_element_by_xpath("//td[contains(@class, 'day text-center')]/a[contains(@data-title, '" + date + "')]").click()
                    sleep(3)
                except Exception:
                    error_flg = True

            for elem_img in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/img[@src='path']"):
                if count == 0:
                    print(date + 'の予定を取得しています。')
                    assign_list.append('\n--------------------------------------------\n')
                    deadline_list.append(date + 'の予定')
                    url_list.append('--------------------------------------------')
                    for elem_h3 in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/h3"):
                        assign_list.append(elem_h3.text)
                    for elem_span in driver.find_elements_by_xpath("//div/span[@class='date pull-xs-right m-r-1']"):
                        deadline_list.append(elem_span.text)
                    for elem_a in driver.find_elements_by_xpath("//div[contains(@class, 'description card-block calendar_event')]/a"):
                        url_list.append(elem_a.get_attribute('href'))
                count += 1
            if count > 0:
                assign += 1

            if count == 0:
                for elem_img in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/img[@src='path']"):
                    if count == 0:
                        print(date + 'の予定を取得しています。')
                        assign_list.append('\n--------------------------------------------\n')
                        deadline_list.append(date + 'の予定')
                        url_list.append('--------------------------------------------')
                        for elem_h3 in driver.find_elements_by_xpath("//div[@class='box card-header clearfix p-y-1']/h3"):
                            assign_list.append(elem_h3.text)
                        for elem_span in driver.find_elements_by_xpath("//div/span[@class='date pull-xs-right m-r-1']"):
                            deadline_list.append(elem_span.text)
                        for elem_a in driver.find_elements_by_xpath("//div[contains(@class, 'description card-block calendar_event')]/a"):
                            url_list.append(elem_a.get_attribute('href'))
                    count += 1
                sub += 1

            if error_flg is False:
                tmp += count
                count = 0

            try:
                calendar_button = driver.find_element_by_link_text('カレンダー')
                calendar_button.click()
                sleep(3)
            except Exception:
                print(date + 'の予定はありませんでした。\n')

            if error_flg is True:
                assign_list.append('\n')
                deadline_list.append(date + 'の予定はありません。')
                url_list.append('\n')

            DateTime = DateTime + datetime.timedelta(days=1)
            error_flg = False
           
    #print(assign_list)
    #print(deadline_list)
    #print(url_list)
    

    #if mode == '1':
    #    print('*課題の数 = ', end = '')
    #    print(assign)
    #if mode == '2':
    #    tmp = tmp - sub
    #    print('*課題の数 = ', end = '')
    #    print(tmp)
    
    # Arduinoへ情報を送信
    if assign == 0:
        ser.write(b"0")
        time.sleep(1)
    if assign == 1:
        ser.write(b"1")
        time.sleep(1)
    if assign == 2:
        ser.write(b"2")
        time.sleep(1)
    if assign > 2:
        ser.write(b"3")
        time.sleep(1)

    print('提出課題情報の取得完了')

    # 片付け
    ser.close()
    driver.close()

def send_mail():
    outlook = win32com.client.Dispatch("Outlook.Application")
    sleep(1)
    mail = outlook.CreateItem(0)

    if mode == '1':
        # メールの内容
        mail.bodyFormat = 1
        mail.to = MailAccount
        mail.subject = '今日の予定のお知らせ'
        body = '今日の予定についてお知らせします。\n今日の予定は以下のようになっています。\n=========================================\n'
        for (assigns, deadline, urls) in zip(assign_list, deadline_list, url_list):
            body += assigns + '    ' + deadline + '\n' + urls + '\n'
        body += '=========================================\n'
        mail.body = body
    if mode == '2':
        # メールの内容
        mail.bodyFormat = 1
        mail.to = MailAccount
        mail.subject = '今日,明日,明後日の予定のお知らせ'
        body = '今日,明日,明後日の予定についてお知らせします。\n3日間の予定は以下のようになっています。\n=========================================\n'
        for (assigns, deadline, urls) in zip(assign_list, deadline_list, url_list):
            body += assigns + '    ' + deadline + '\n' + urls + '\n'
        body += '=========================================\n'
        mail.body = body

    #mail.Display(True)  # 自動作成したメールの確認
    mail.Send()  # 確認せずに送信する場合

    print('メール送信完了')
    assign_list.clear()
    deadline_list.clear()
    url_list.clear()

def job():
    get_information()
    send_mail()

print('*モード選択*')
print('1：今日の予定を調査します。')
print('2：今日から明後日までの3日間の予定を調査します。')
global mode
assign_list = []
deadline_list = []
url_list = []
mode = input('モード( 1 or 2 )を入力してください。\n')
clock = input('実行したい時刻を半角で入力してください。( 例：08:30 )\n')
#schedule.every(1/60).minutes.do(job)
schedule.every().day.at(clock).do(job)

while True:
    schedule.run_pending()
    time.sleep(1)
