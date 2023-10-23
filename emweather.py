# -*- coding: utf-8 -*-
# ======================================
# 
# 電子マネー管理システム　フレームワーク
# 天候情報取得モジュール
#
# [環境]
#   Python 3.10.8
#   python3-tk Ver3.8.10
#   VSCode 1.64
#   <拡張>
#     |- Python  V2021.12
#     |- Pylance V2021.12
#
# [更新履歴]
#   2022/3/4 新規作成
# ======================================
import datetime
import csv
import urllib.request
import sys
sys.path.append("lib.bs4")
from bs4 import BeautifulSoup

def str2float(weather_data):
    try:
        return float(weather_data)
    except:
        return 0

def scraping(url,year,month,prec,block):

    # 気象データのページを取得
    html = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(html,features='lxml')
    #soup = BeautifulSoup(html, "html.parser")
    #print(soup)
    #print(type(soup.string))
    trs = soup.find("table", { "class" : "data2_s" })
    #print(type(trs))
    data_list = []
    data_list_per_hour = []

    # table の中身を取得
    for tr in trs.findAll('tr')[4:]:
        tds = tr.findAll('td')

        if tds[0].string == None:
            break;

        data_list.append(str2float(str(tds[0].string)))
        data_list.append(str(tds[19].string))
        data_list.append(str(tds[20].string))
        #data_list.append(tds[2].string)
        #data_list.append(str2float(tds[3].string))
        #data_list.append(str2float(tds[4].string))
        #data_list.append(str2float(tds[5].string))
        #data_list.append(str2float(tds[6].string))
        data_list.append(str2float(str(tds[7].string)))
        data_list.append(str2float(str(tds[8].string)))
        #data_list.append(str2float(tds[9].string))
        #data_list.append(str2float(tds[10].string))
        #data_list.append(str2float(tds[11].string))
        #data_list.append(str2float(tds[12].string))
        #data_list.append(str2float(tds[13].string))
        data_list.append(year)
        data_list.append(month)
        data_list.append(prec)
        data_list.append(block)
        
        data_list_per_hour.append(data_list)

        data_list = []

    return data_list_per_hour

def create_wether_csv(prec,block,year,month,output_file): 
    # CSV 出力先
    # output_file = jp.weather_json()    

    # データ取得開始・終了日
    start_date = datetime.date(int(year), int(month), 1)
    #end_date   = datetime.date(2021, 12, 24)
    viewcd ="p1"

    # CSV の列
    fields = ["日","昼(06:00-18:00)", "夜(18:00-翌日06:00)","最高気温","最低気温","年","月","prec","block"] 
    

    #with open(output_file, 'w') as f:
    with open(output_file, 'w') as f:
        writer = csv.writer(f, lineterminator='\n')
        # if head_flg == '1':
        #     writer.writerow(fields)
        #ヘッダーを書き込む
        writer.writerow(fields)
        date = start_date
        # 対象url
        url = "http://www.data.jma.go.jp/obd/stats/etrn/view/daily_s1.php?" \
                "prec_no=%s&block_no=%s&year=%d&month=%d&day=%d&view=%s"%(prec,block,date.year, date.month, date.day,viewcd)
            
        data_per_day = scraping(url,date.year,date.month,prec,block)

        for dpd in data_per_day:
            writer.writerow(dpd)
def weather_list_get(prec,block,year,month): 
    
    start_date = datetime.date(int(year), int(month), 1)
    viewcd ="p1"

    date = start_date
    # 対象url
    url = "http://www.data.jma.go.jp/obd/stats/etrn/view/daily_s1.php?" \
            "prec_no=%s&block_no=%s&year=%d&month=%d&day=%d&view=%s"%(prec,block,date.year, date.month, date.day,viewcd)
        
    data_per_day = scraping(url,date.year,date.month,prec,block)
    
    data_list = []
    for dpd in data_per_day:
        data_list.append(dpd)
    
    return data_list
            
#
# 指定された地点、日付の天候を検索
#
def get_wether(prec,block,year,month,day): 
    
    # データ取得開始
    sd = datetime.date(int(year), int(month), 1)
    # 検索キーワード
    viewcd ="p1"

    # 戻り値
    # fields = ["日","昼(06:00-18:00)", "夜(18:00-翌日06:00)","最高気温","最低気温","年","月","prec","block"] 
       
    # 対象url
    url = "http://www.data.jma.go.jp/obd/stats/etrn/view/daily_s1.php?" \
            "prec_no=%s&block_no=%s&year=%d&month=%d&day=%d&view=%s"%(prec, block, sd.year, sd.month, sd.day, viewcd)
        
    data_per_day = scraping(url,year,month,prec,block)

    return data_per_day[day - 1]
       

