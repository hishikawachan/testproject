# -*- coding: utf-8 -*-
# ======================================
# 
# 電子マネー管理システム
# 期間帳票出力メインモジュール　
# [環境]
#   Python 3.10.8
#
#   <拡張>\\
#     |- Python  V2021.12
#     |- Pylance V2021.12
#
# [更新履歴]
#   2023/8/21  新規作成
# ======================================
from datetime import datetime
import datetime
#import calendar
import yaml 
from emoneydbclass import DataBaseClass
from emoneydbreportclass import dbReport

#################################################################
# DB制御classに渡すパラメータ
#################################################################
parm_data = []

#################################################################
# メイン
#################################################################
if __name__ == "__main__":
    #
    # Configファイルから入力ファイル名等抽出    
    #
    # パラメタ
    # [i]:dbip
    # [1]:dbname
    # [2]:dbport
    # [3]:dbuser
    # [4]:dbpassword
    # [5]:処理対象開始日
    # [6]:処理対象終了日
    # [7]:処理区分(1:売上明細 2:TOAMAS 3:ヤマトフィナンシャルデータ)
    # [8]:対象会社コード
    # [9]:帳票ファイル出力先ディレクトリ
    # 
    # # 対象データリスト
    
    dt_now = datetime.datetime.now()
    print('期間帳票出力処理開始：',dt_now) 
    
    input_symd = input('処理開始日を入力(yyyyMMdd 99999999は終了):') 
    input_eymd = input('処理終了日を入力(yyyyMMdd 99999999は終了):') 
    input_comcd =  input('会社コードを入力(9999999 9999999は終了):')    
    
    if (input_symd != '99999999' and input_eymd  != '99999999' and input_comcd != '9999999'):
        s_year = int(input_symd[0:4])
        s_month = int(input_symd[4:6]) 
        e_year = int(input_eymd[0:4])
        e_month = int(input_eymd[4:6]) 
        s_day = int(input_symd[6:8]) 
        e_day = int(input_eymd[6:8]) 
        #前月の初日と最終日を求める
        # d = datetime.datetime.today()
        # today = d.date()        
        # month = today.month - 1
        # year = today.year
        # s_day = 1
        # if month <= 0:
        #     month = 12
        #     year = year - 1 
        # e_day = calendar.monthrange(year, month)[1]
        
        # 基本情報取得
        with open('C:/emoney/emoney.yaml','r+',encoding="utf-8") as ry:
            config_yaml = yaml.safe_load(ry)
            dbip = config_yaml['dbip']
            dbmarianame = config_yaml['dbmarianame']
            dbport = config_yaml['dbport']        
            dbuser = config_yaml['dbuser']
            dbpw = config_yaml['dbpw']
            parm_data = []
            parm_data.append(dbip)
            parm_data.append(dbmarianame)
            parm_data.append(dbport)        
            parm_data.append(dbuser)
            parm_data.append(dbpw)           
            
            file_path = config_yaml['dir_filepath']
            parm_data.append(file_path)
        
        #データベース操作クラス初期化
        resdb = DataBaseClass(parm_data)   
            
        #全会社データ取得
        ret_rows = resdb.company_data_allget()
        del resdb
            
        ###############################
        # 会社データを順次読み月次帳票処理
        ###############################
        for i in range(0,len(ret_rows)):
            if ret_rows[i][0] == input_comcd: #処理対象の会社コードかチェック
                print('対象会社  :',ret_rows[i][1])
                parm_data = []
                parm_data.append(dbip)
                parm_data.append(dbmarianame)
                parm_data.append(dbport)        
                parm_data.append(dbuser)
                parm_data.append(dbpw)
                start_date = datetime.date(s_year,s_month,s_day)
                end_date = datetime.date(e_year,e_month,e_day)
                parm_data.append(start_date)
                parm_data.append(end_date)
                parm_data.append(ret_rows[i][9])
                parm_data.append(ret_rows[i][0])
                parm_data.append(file_path)
                
                #帳票作成クラス初期化
                resdbrep = dbReport(parm_data)
                res_cont = resdbrep.main()
                del resdbrep 
            
            i += 1
        
    dt_now = datetime.datetime.now()
    print('期間帳票出力処理終了：',dt_now) 
        
       
        
        
        
        
        
        
        
        
        
        
