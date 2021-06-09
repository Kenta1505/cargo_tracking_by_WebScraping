from selenium import webdriver # seleniumというブラウザを自動操作するライブラリ
from bs4 import BeautifulSoup # BeautifulSoupでスクレイピングを行う
import excel_update # 先ほど作ったエクセル処理のpyファイルを読み込み
from time import sleep

def web_scraping():
    track_no=excel_update.excel_update() # エクセルから取得したセルの値を変数に代入
    print(track_no) # print関数で変数に値が正しく代入されていることがわかる
    
    driver=webdriver.Chrome('chromedriver_win32\chromedriver') # seleniumを使って、Google Chromeを起動
    driver.get('https://toi.kuronekoyamato.co.jp/cgi-bin/tneko') # 対象のページへ移動
    
    search_bar=driver.find_element_by_name('number01') # HTML/CSSのタグを頼りにページ内の情報を取得していく
    search_bar.send_keys(track_no) # ここでは、先ほど取得した追跡番号を検索バーに入力する処理を行っている
    
    button_1=driver.find_element_by_name('sch') # お問い合わせボタンを探し出し、クリック
    button_1.click()
    sleep(2) # ページの読み込みを待つため、一応2秒待機
    
    html=driver.page_source # 読み込んだページの情報を取得し、変数htmlに代入
    
    soup=BeautifulSoup(html, 'lxml') # BeautifulSoupを使って、変数htmlに入っているページ情報を読み込む
    data=soup.find_all(True, {'class':'meisaibase'}) # 読み込んだページ情報の中から、欲しいデータが含まれているHTML/CSSのタグを見つけ、情報を抽出
    print(data) # print関数で情報が正しく取れているか確認しながら行うとわかりやすい
    
    for x in range(len(data)):
        arrival=data[int(x)].find_all('td')
        print(arrival) # 情報が正しく取得できているか確認
        
    date=arrival[1].get_text()
    arrival_date=date.split('　')[0] # ここまで来たらあとは取得したテキスト情報をキレイに整えていく。ここは到着予定日の情報
    arrival_time=date.split('　')[1] # こっちは、到着予定時刻
    
    driver.close() # 最初にseleniumで立ち上げたブラウザを終了する
    
    excel_update.excel_renew(arrival_date, arrival_time) # excel_updateの中のエクセルに値を入力していく関数の引数に今回の結果を代入して、処理を実行
    
    
    

if __name__ == '__main__': #作成した関数を実行するための処理
    web_scraping()
