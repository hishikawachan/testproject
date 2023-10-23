# -*- coding: utf-8 -*-
# ======================================
# 
# 電子マネー管理システム
# 全体制御メインモジュール
# 処理対象の会社に対してDBへデータ出力
# 売上集計帳票を出力
# [環境]
#   Python 3.10.8
#   VSCode 1.64
#   <拡張>
#     |- Python  V2021.12
#     |- Pylance V2021.12
#
# [更新履歴]
#   2023/10/21  新規作成
# ======================================
from datetime import datetime
import datetime
import csv
import yaml
import os 
from emdbclass import DataBaseClass
from emreportclass import dbReport
from emdbedit import dbEditor
#################################################################
# 共通パラメータ
#################################################################
parm_data = []

#################################################################
# メイン
#################################################################
if __name__ == "__main__":
    #
    # Configファイルから入力ファイル名等抽出    
    #
    # 共通パラメタ
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
    # 対象データリスト
    
    dt_now = datetime.datetime.now()
    print('全体処理開始：',dt_now) 
    
    # 基本情報取得
    with open('C:/Users/hishi/em/testproject/src/em.yaml','r+',encoding="utf-8") as ry:
        config_yaml = yaml.safe_load(ry)
        dbip = config_yaml['dbip']
        dbmarianame = config_yaml['dbmarianame']
        dbport = config_yaml['dbport']        
        dbuser = config_yaml['dbuser']
        dbpw = config_yaml['dbpw']
        parm_data.append(dbip)
        parm_data.append(dbmarianame)
        parm_data.append(dbport)        
        parm_data.append(dbuser)
        parm_data.append(dbpw)
        
        file_path = config_yaml['dir_filepath']
        parm_data.append(file_path)
        
        d = datetime.datetime.today()
        today = d.date()        
        #データベース操作クラス初期化
        resdb = DataBaseClass(parm_data) 
        #データベースバックアップ実行
        print('データベースバックアップ(処理前)開始：',datetime.datetime.now())         
        ret = resdb.database_backup()        
        print('データベースバックアップ(処理前)終了：',datetime.datetime.now())
         
        #会社データ全件取得
        ret_rows = resdb.company_data_allget()
        del resdb
        
        ########################################
        #
        # 会社データ毎の処理
        # 会社DBをシーケンスに読み、処理対象の会社を検知したらDB書込み及び帳票出力を行う
        # 会社データ　処理予定日 <= 処理日
        #
        ########################################        
        for i in range(0,len(ret_rows)):
            # 今日の日付と登録された処理予定日の比較
            t_date =  ret_rows[i][4]
            td = today - t_date
            # 処理予定日<=今日なら処理対象とする
            if td.days >= 0:
                #debug
                print('処理対象会社名 :',ret_rows[i][1])                
                print('データリスト化開始 :',datetime.datetime.now())
                parm_data = []
                #共通パラメータセット
                parm_data.append(dbip)
                parm_data.append(dbmarianame)
                parm_data.append(dbport)        
                parm_data.append(dbuser)
                parm_data.append(dbpw)
                parm_data.append(ret_rows[i][7])  #更新開始日セット
                parm_data.append(ret_rows[i][8])  #更新終了日セット
                parm_data.append(ret_rows[i][9])  #処理区分セット
                parm_data.append(ret_rows[i][0])  #対象会社コード
                parm_data.append(file_path)       #帳票出力先ディレクトリ
                
                # 処理対象会社の売上明細又はTOAMASファイルのデータを全件リスト化
                proc_flg = '0'
                file_kbn = '0'
                input_rows1 = []
                input_rows2 = []                
                f_name = ret_rows[i][10]
                file_name = f_name.strip()
                input_filepath = os.path.join(file_path,file_name)  
                if ret_rows[i][9] == '1':
                    file_kbn = '1'
                    if  proc_flg == '0': #売上明細ファイルを全件読み込みリスト化(初回のみ)           
                        with open(input_filepath, encoding = 'shift-jis') as f:
                            reader = csv.reader(f)
                            for row in reader :
                                input_rows1 = [row for row in reader]
                            proc_flg = '1'
                else:
                    if ret_rows[i][9] == '2': #指定TOAMASファイル全件読み込みリスト化(対象会社毎)
                        file_kbn = '2'  
                        with open(input_filepath, encoding = 'UTF-8') as f:
                            reader = csv.reader(f)
                            for row in reader :
                                input_rows2 = [row for row in reader]      
                #debug
                print('データリスト化終了 :',datetime.datetime.now())
                ########################################
                #
                # 売上データの抽出
                # 全件読み込んだ入力データ（リスト）から処理対象を抽出
                #
                ########################################       
                output_rows = []
                input_rows = []
                now_placecd = ''
                x = 0
                #データベース操作クラス初期化
                resdb = DataBaseClass(parm_data) 

                if file_kbn == '1':
                    input_rows.append(input_rows1) 
                if file_kbn == '2':
                    input_rows.append(input_rows2) 
                #読み込んだデータがあれば
                if len(input_rows) > 0:
                    #debug
                    print('全件リストからデータ抽出開始 :',datetime.datetime.now())
                    
                    for x in range(0,len(input_rows)):
                        if file_kbn == '1': #売上照会ファイル
                            if now_placecd != input_rows[x][2]:
                                #設置場所コードが変わったら　会社コード検索    
                                place_row = resdb.place_data_get(input_rows[x][2]) #設置場所コードで会社コード検索
                                now_placecd = input_rows[x][2]
                            str_placecd = str(place_row[0][3])[0:7]#改行コード外し
                            if str_placecd == ret_rows[i][0]: #会社コード一致
                                input_year = str(input_rows[x][18])[0:4]
                                input_month = str(input_rows[x][18])[4:6]
                                input_day = str(input_rows[x][18])[6:8]
                                terget_date = datetime.date(int(input_year), int(input_month), int(input_day))
                                s_date = ret_rows[i][7] 
                                td = s_date - terget_date #対象開始日との比較
                                if td.days <= 0:
                                    e_date = ret_rows[i][8]
                                    td = e_date - terget_date #対象終了日との比較 
                                    if td.days >= 0:
                                        output_rows.append(input_rows[x]) #対象データ保管
                        else:
                            if file_kbn == '2': #TOAMASファイル
                                if ret_rows[i][0] == str(input_rows[x][8])[0:7]: #会社コード一致（資産管理コードの上位7桁）
                                    datetime_list = input_rows[x][0].split()
                                    datetime_date = datetime_list[0].split('-')
                                    datetime_time = datetime_list[1].split(':')       
                        
                                    terget_date = datetime.date(int(datetime_date[0]),int(datetime_date[1]),int(datetime_date[2])) 
                                    s_date = ret_rows[i][7] 
                                    td = s_date - terget_date #対象開始日との比較
                                    if td.days <= 0:
                                        e_date = ret_rows[i][8]
                                        td = e_date - terget_date #対象終了日との比較 
                                        if td.days >= 0:
                                            output_rows.append(input_rows[x]) #対象データ保管
                                                                       
                        x += 1 
                    del resdb 
                    
                    #debug  
                    print('全件リストからデータ抽出終了 :',datetime.datetime.now())
                    
                    if len(output_rows) > 0: #対象データがあればDB書込み処理
                        #debug  
                        print('抽出データからDB出力処理開始 :',datetime.datetime.now())

                        # データ編集/出力クラス初期化
                        res_ed = dbEditor(parm_data,output_rows)
                        # 売上照会ファイル編集・書込み
                        if file_kbn == '1':
                            print('売上照会データ処理開始')
                            res_cont = res_ed.uriage_edit()
                        # TOAMASファイル編集・書込み
                        if file_kbn  == '2':
                            print('TOAMASデータ処理開始')
                            res_cont = res_ed.toamas_edit()

                        # ヤマトフィナンシャルファイル編集・書込み
                        # KGS故障時データがアップロードできていない場合に使用
                        # if parm_data[7]  == '3':
                        #     print('ヤマトフィナンシャルデータ処理開始')
                        #     res_cont = res_ed.yamato_edit(parm_data)
                            
                        del res_ed                    
                        #
                        # res_cont[0] : 結果ステータス
                        # res_cont[1] : DB 出力件数
                        # res_cont[2] : DB 出力不可件数
                        # res_cont[3] : DB 編集日付
                        #
                        #debug  
                        print('抽出データからDB出力処理終了 :',datetime.datetime.now())

                        if res_cont[0] != 0:
                            print('データベースの出力失敗 status: ',res_cont[0])
                        else:                        
                            #帳票作成クラス初期化~帳票出力
                            #debug  
                            print('帳票出力処理開始 :',datetime.datetime.now())
                            resdbrep = dbReport(parm_data)
                            res_cont = resdbrep.main()
                            del resdbrep
                            #対象会社データ次回処理日の設定
                            resdb = DataBaseClass(parm_data) 
                            res_row = resdb.company_updateday_update(ret_rows[i][0])
                            del resdb
                            #debug  
                            print('帳票出力処理終了 :',datetime.datetime.now())  
            i += 1
        
        print('データベースバックアップ開始(処理後) :',datetime.datetime.now()) 
        resdb = DataBaseClass(parm_data) 
        ret = resdb.database_backup()
        print('データベースバックアップ終了(処理後) :',datetime.datetime.now()) 
        del resdb  
        
        print('処理終了：',datetime.datetime.now())            