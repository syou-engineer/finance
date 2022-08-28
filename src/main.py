from dotenv import load_dotenv
import math
import time
import datetime
import os
from dateutil.relativedelta import relativedelta

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

import chart

# 権利落ち日
ex_rights_date = [
    datetime.datetime(2022,1,27),
    datetime.datetime(2022,2,24),
    datetime.datetime(2022,3,29),
    datetime.datetime(2022,4,26),
    datetime.datetime(2022,5,27),
    datetime.datetime(2022,6,28),
    datetime.datetime(2022,7,27),
    datetime.datetime(2022,8,29),
    datetime.datetime(2022,9,28),
    datetime.datetime(2022,10,27),
    datetime.datetime(2022,11,28),
    datetime.datetime(2022,12,28),
    ]

def connect_sbi(ACCOUNT, PASSWORD, name):
    options = Options()
    # ヘッドレスモード(chromeを表示させないモード)
    options.add_argument('--headless')
    # driver = webdriver.Chrome(options=options, executable_path='chromedriver')
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

    # SBI証券のトップ画面を開く
    driver.get('https://www.sbisec.co.jp/ETGate')    

    # タイムアウトを10秒に設定
    wait = WebDriverWait(driver, 10)
    # すべての要素が表示されるまで待機
    element = wait.until(expected_conditions.visibility_of_all_elements_located)

    # ユーザーIDとパスワード
    input_user_id = driver.find_element_by_name('user_id')
    input_user_id.send_keys(ACCOUNT)
    input_user_password = driver.find_element_by_name('user_password')
    input_user_password.send_keys(PASSWORD)
    
    # ログインボタンをクリック
    driver.find_element_by_name('ACT_login').click()

    return driver


def get_ja_data(driver):
    # 遷移するまで待つ
    time.sleep(4)

    # ポートフォリオの画面に遷移
    driver.find_element_by_xpath('//*[@id="link02M"]/ul/li[1]/a/img').click()

    driver.find_element_by_xpath('/html/body/div[3]/div/table/tbody/tr/td/table[2]/tbody/tr/td/a[3]').click()

    time.sleep(4)

    # 文字コードをUTF-8に変換
    html = driver.page_source.encode('utf-8')

    # BeautifulSoupでパース
    soup = BeautifulSoup(html, "html.parser")

    # 株式
    table_data = soup.find_all("table", bgcolor="#9fbf99", cellpadding="4", cellspacing="1", width="100%") 

    # 株式（現物/特定預り）
    df_stock_specific = pd.read_html(str(table_data), header=0)[0]
    df_stock_specific = format_data(df_stock_specific, '株式（現物/特定預り）', '上場ＴＰＸ')

    # 株式（現物/NISA預り）
    df_stock_fund_nisa = pd.read_html(str(table_data), header=0)[1]
    df_stock_fund_nisa = format_data(df_stock_fund_nisa, '株式(現物/NISA預り)', '上場ＴＰＸ')

    # 株式（信用）
    # print(pd.read_html(str(table_data), header=0)[2])
    # df_stock_fund_nisa = pd.read_html(str(table_data), header=0)[2]
    # df_stock_fund_nisa = format_data_Shinyou(df_stock_fund_nisa, '株式（信用）', '上場ＴＰＸ')

    # 投資信託（金額/特定預り）
    # df_fund_specific = pd.read_html(str(table_data), header=0)[3]
    # df_fund_specific = format_data(df_fund_specific, '投資信託（金額/特定預り）', '')

    # # 投資信託（金額/NISA預り）
    # df_fund_nisa = pd.read_html(str(table_data), header=0)[4]
    # df_fund_nisa = format_data(df_fund_nisa, '投資信託（金額/NISA預り）', '')

    # # 投資信託（金額/つみたてNISA預り）
    # df_fund_nisa_tsumitate = pd.read_html(str(table_data), header=0)[5]
    # df_fund_nisa_tsumitate = format_data(df_fund_nisa_tsumitate, '投資信託（金額/つみたてNISA預り）', '')

    # 結合
    # df_ja_result = pd.concat([df_stock_specific, df_stock_fund_nisa, df_fund_specific, df_fund_nisa, df_fund_nisa_tsumitate])
    df_ja_result = pd.concat([df_stock_specific, df_stock_fund_nisa])
    df_ja_result['date'] = datetime.date.today()

    return df_ja_result

def get_ja_data_split(driver):
    # 遷移するまで待つ
    time.sleep(4)

    # 分割表示画面に切り替える    
    driver.find_element_by_xpath('/html/body/div[3]/div/table/tbody/tr/td/table[2]/tbody/tr/td/a[3]').click()
    
    time.sleep(4)
    # 分割表示画面に切り替える
    driver.find_element(by=By.XPATH, value='/html/body/div[3]/div/table/tbody/tr/td/table[4]/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/a').click()
    time.sleep(4)

    # 文字コードをUTF-8に変換
    html = driver.page_source.encode('utf-8')

    # BeautifulSoupでパース
    soup = BeautifulSoup(html, "html.parser")

    # 株式
    table_data = soup.find_all("table", bgcolor="#9fbf99", cellpadding="4", cellspacing="1", width="100%") 

    # 株式（現物/特定預り）
    df_stock_specific = pd.read_html(str(table_data), header=0)[0]
    df_stock_specific = format_data_split(df_stock_specific, '株式（現物/特定預り）', '上場ＴＰＸ')

    # 株式（現物/NISA預り）
    df_stock_fund_nisa = pd.read_html(str(table_data), header=0)[1]
    df_stock_fund_nisa = format_data_split(df_stock_fund_nisa, '株式(現物/NISA預り)', '上場ＴＰＸ')

    # 株式（信用）
    # print(pd.read_html(str(table_data), header=0)[2])
    # df_stock_fund_nisa = pd.read_html(str(table_data), header=0)[2]
    # df_stock_fund_nisa = format_data_Shinyou(df_stock_fund_nisa, '株式（信用）', '上場ＴＰＸ')

    # 投資信託（金額/特定預り）
    # df_fund_specific = pd.read_html(str(table_data), header=0)[3]
    # df_fund_specific = format_data(df_fund_specific, '投資信託（金額/特定預り）', '')

    # # 投資信託（金額/NISA預り）
    # df_fund_nisa = pd.read_html(str(table_data), header=0)[4]
    # df_fund_nisa = format_data(df_fund_nisa, '投資信託（金額/NISA預り）', '')

    # # 投資信託（金額/つみたてNISA預り）
    # df_fund_nisa_tsumitate = pd.read_html(str(table_data), header=0)[5]
    # df_fund_nisa_tsumitate = format_data(df_fund_nisa_tsumitate, '投資信託（金額/つみたてNISA預り）', '')

    # 結合
    # df_ja_result = pd.concat([df_stock_specific, df_stock_fund_nisa, df_fund_specific, df_fund_nisa, df_fund_nisa_tsumitate])
    df_ja_result = pd.concat([df_stock_specific, df_stock_fund_nisa])
    df_ja_result['date'] = datetime.date.today()

    return df_ja_result


def format_data(df_data, category, fund):
    # 必要な列のみ抽出
    df_data = df_data.loc[:, ['銘柄（コード）', '評価額', '損益', '損益（％）']]

    df_data['カテゴリー'] = category
    if fund != '':
        df_data['ファンド名'] = fund

    return df_data

def format_data_split(df_data, category, fund):
    # 必要な列のみ抽出
    df_data = df_data.loc[:, ['銘柄（コード）', '買付日', '数量']]

    df_data['カテゴリー'] = category
    if fund != '':
        df_data['ファンド名'] = fund

    return df_data

def format_data_Shinyou(df_data, category, fund):
    # 必要な列のみ抽出
    df_data = df_data.loc[:, ['建代金', '損益', '損益（％）']]

    df_data['カテゴリー'] = category
    if fund != '':
        df_data['ファンド名'] = fund

    return df_data

def export_data(df_result, name):
    df_result.to_csv(name + '_result_{0:%Y%m%d}.txt'.format(datetime.date.today()))


def devided_collection(driver, df: pd.DataFrame):
    # 銘柄が重複してる場合（Nisa/特定）は取り除く
    sr = df['銘柄（コード）'].drop_duplicates()

    
    kimatu_sr = list()
    chukan_sr = list()
    code_sr = list()
    kimatu_date_sr = list()
    chukan_date_sr = list()
    for d in sr:
        code_sr.append(str(d))

        code = d[:4]
        path = f'https://irbank.net/{code}'
        # 別タブを開いて配当金データを取得しにいく
        driver.execute_script("window.open('about:blank', 'tab2');")
        driver.switch_to.window("tab2")
        driver.get(path)
        
        # タイムアウトを10秒に設定
        wait = WebDriverWait(driver, 10)
        # すべての要素が表示されるまで待機
        element = wait.until(expected_conditions.visibility_of_all_elements_located)
        # 配当推移を開く
        devidec_element = driver.find_element(by=By.XPATH, value='//*[@id="c_Link"]/div/div/div/ul[2]/li[7]/a')
        driver.execute_script('arguments[0].click();', devidec_element)

                # タイムアウトを10秒に設定
        wait = WebDriverWait(driver, 10)
        # すべての要素が表示されるまで待機
        element = wait.until(expected_conditions.visibility_of_all_elements_located)

            # 文字コードをUTF-8に変換
        html = driver.page_source.encode('utf-8')

    # BeautifulSoupでパース
        soup = BeautifulSoup(html, "html.parser")

        # 株式
        table_data = soup.find("table", class_="cs") 
        df_devidend = pd.read_html(str(table_data), header=0)[0]

        # year_month = df_devidend['年度'].iloc[-1].strftime('%Y年%m月')
        last_index = df_devidend.iloc[-1]
        if check_dividend_record_date(last_index['年度']):
            # 年度またぎの場合は中間と期末を取得する必要がある
            year_month = df_devidend.tail(2)
            kimatu = year_month.iloc[0]['期末']
            kimatu_sr.append(str(kimatu))
            chukan = year_month.iloc[1]['中間']
            chukan_sr.append(str(chukan))
            kimatu_date_sr = date_parse(year_month.iloc[0]['年度'])
            chukan_date_sr = date_parse(year_month.iloc[0]['年度']) + relativedelta(months=6)
            print('期末：' + str(kimatu) + ' 中間：' +  str(chukan))
        else:
            year_month = df_devidend.tail(1)
            kimatu = year_month.iloc[0]['期末']
            if '中間' in year_month.columns:
                chukan = year_month.iloc[0]['中間']
            else:
                chukan = None
            kimatu_sr.append(str(kimatu))
            chukan_sr.append(str(chukan))
            kimatu_date_sr = date_parse(year_month.iloc[0]['年度'])
            chukan_date_sr = date_parse(year_month.iloc[0]['年度']) + relativedelta(months=6)
        # df_devidend = df_devidend.reindex(columns=['年度', '区分', '中間', '期末', '合計', "配当利回り", "備考"])
        # df_devidend = df_devidend.loc[:, [('年度', '区分', '中間', '期末', '合計')]]
    
    return pd.DataFrame(
        data={
            '銘柄（コード）': code_sr,
            '中間': chukan_sr,
            '期末': kimatu_sr,
            '中間権利確定月': chukan_date_sr,
            '期末権利確定月': kimatu_date_sr,            
        }
    )

def date_parse(date):
    parse_date = datetime.datetime.strptime(date, '%Y年%m月')
    return parse_date

def check_dividend_record_date(date):
    """配当基準日をチェック、True:年度跨ぎ（1〜3月）"""
    month = date_parse(date).month
    if month <= 3:
        return True
    else:
        return False

# interim_dividend, termend_dividend, termend_dividend_date, termend_dividend_date, stock_num
def dividend_calc(df: pd.DataFrame, interim_dividend, termend_dividend, interim_dividend_date, termend_dividend_date):
    """配当金の計算"""
    datas = {}
    result_interim_dividend = 0
    result_termend_dividend = 0

    for index, colmun in enumerate(df.columns, 0):
        datas[colmun] = index

    for row in df.values:
        # something to do
        print(row)
        buy_date = datetime.datetime.strptime(row[datas['買付日']], '%y/%m/%d')
        stock_num = int(row[datas['数量']])
        # 中間配当の月
        # if interim_dividend_date.get('中間権利確定月'):
        interim_month = int(interim_dividend_date.dt.month.iloc[-1])
        # 中間配当の権利落ち日を取得
        interim_ex_date = ex_rights_date[interim_month - 1]
        
        if buy_date <= interim_ex_date:
            result_interim_dividend += stock_num * interim_dividend

        # 期末配当の月
        termend_month = int(termend_dividend_date.dt.month.iloc[-1])
        # 期末配当の権利落ち日を取得
        termend_ex_date = ex_rights_date[termend_month - 1]

        if buy_date <= termend_ex_date:
            result_termend_dividend += stock_num * termend_dividend


    # 中間 + 期末配当の合計値
    return int(result_interim_dividend + result_termend_dividend)

if __name__ == '__main__':
    # .envから環境変数を読み込む
    load_dotenv()

    ACCOUNT = os.getenv('ACCOUNT')
    print(ACCOUNT)
    PASSWORD = os.getenv('PASSWORD')
    print(PASSWORD)
    name = 'NAMAE'

    # SBI証券に接続
    driver = connect_sbi(ACCOUNT, PASSWORD, name)
    # ポートフォリオからデータ取得
    df_ja_result = get_ja_data(driver)
    df_ja_result_split = get_ja_data_split(driver)

    # 配当金情報を取得してくる
    df_devided_table = devided_collection(driver, df_ja_result_split)

    # 年間配当金を算出
    price = 0
    for _name, _df in df_ja_result_split.groupby('銘柄（コード）'):
        meigara_price = 0

        query=f'"銘柄（コード）"=="{_name}"'
        print(query)
        # df = df_devided_table.query(query)
        df = df_devided_table[df_devided_table['銘柄（コード）'] == f'{_name}']
        interim_devided = df['中間'].iloc[-1]
        if interim_devided == "None":
            interim_devided = 0
        else:
            if str(termend_devided).isdecimal():
            # math.floorで小数点を切り捨てる
                interim_devided = math.floor(float(interim_devided))
            else:
                # 数値以外のパターン "-" などが存在するため数値変換ミスの場合は0円とする
                interim_devided = 0

        termend_devided = df['期末'].iloc[-1]

        if termend_devided == "None":
            termend_devided = 0
        else :
            # math.floorで小数点を切り捨てる
            if str(termend_devided).isdecimal():
                termend_devided = math.floor(float(termend_devided))
            else:
                # 数値以外のパターン "-" などが存在するため数値変換ミスの場合は0円とする
                termend_devided = 0

        interim_devided_date = pd.to_datetime(df['中間権利確定月'], format = '%d/%m/%Y')
        termend_devided_date = pd.to_datetime(df['期末権利確定月'], format = '%d/%m/%Y')
        print(f'中間権利確定月：{interim_devided_date}')
        print(f'期末権利確定月{termend_devided_date}')
        meigara_price += dividend_calc(
            _df, 
            interim_devided, 
            termend_devided, 
            interim_devided_date,
            termend_devided_date
            # datetime.datetime.strptime(str(interim_devided_date), "%Y-%m-%d"),
            # datetime.datetime.strptime(str(termend_devided_date), "%Y-%m-%d")
            )
        price += meigara_price
        print(f'配当金の合計：{price}円')
        print(f'{_name}：配当金の合計：{meigara_price}円')

    export_data(df_ja_result, name)
    export_data(df_ja_result_split, "NAMAE_SPLIT")
    export_data(df_devided_table, "NAMAE_DEVIDED")

    
    # チャートに描画
    chart.draw("日本株 各銘柄保有割合", df_ja_result['評価額'], df_ja_result['銘柄（コード）'])
    
    # ブラウザーを閉じる
    driver.quit() 