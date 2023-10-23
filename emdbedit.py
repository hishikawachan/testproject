# -*- coding: utf-8 -*-
# ======================================
# 
# 電子マネー管理システム
# 売上集計csvファイル入力
# MariaDB データベース出力データ編集　
# [環境]
#   Python 3.10.8
#   VSCode 1.64
#   <拡張>
#     |- Python  V2021.12
#     |- Pylance V2021.12
#
# [更新履歴]
#   2023/2/20  新規作成
# ======================================

from datetime  import datetime
import datetime
from emoneydbclass import DataBaseClass

class dbEditor:
    #####################################
    # 初期化
    #
    # パラメタ
    # データ配列
    #####################################
    def __init__(self,parm,data):
        self.row = data        
        # DB制御クラス初期化               
        self.resdb = DataBaseClass(parm)
        
    #####################################
    # 売上照会ファイル編集
    #####################################
    def uriage_edit(self):
        
        #[2]設置場所番号
        #[3]設置場所名称
        #[12]明細区分番号
        #[13]明細区分名称
        #[14]明細種別番号
        #[15]明細種別名称
        #[17]決済金額
        #[18]決済時刻
        
        #debug
        dt_now = datetime.datetime.now()
        print('売上データ出力処理開始(売上明細)：',dt_now)     
        #
        
        out_count = 0
        in_count = 0
        out_err = 0
        output_list = []
        sum_price = 0
        
        for row in self.row:
            #練習データを排除する
            if int(row[14]) <= 10: #明細種別11以上は練習データ
                in_count += 1
                data_list = []
                date_time = row[18]
                data_list.append(int(date_time[0:4]))#決済年
                data_list.append(int(date_time[4:6]))#決済月
                data_list.append(int(date_time[6:8]))#決済日
                data_list.append(int(date_time[8:10]))#決済時
                data_list.append(int(date_time[10:12]))#決済分
                data_list.append(int(date_time[12:14]))#決済秒
                
                kyear = int(date_time[0:4])#決済年
                kmonth = int(date_time[4:6])#決済月
                kday = int(date_time[6:8])#決済日
                khour = int(date_time[8:10])#決済時
                kminute = int(date_time[10:12])#決済分
                ksecond = int(date_time[12:14])#決済秒            
                
                data_list.append(int(date_time[14:16]))#決済番号
                data_list.append(int(row[2])) #設置場所番号
                #data_list.append(str(row[12])) #明細区分番号
                if row[12] == '00':#景品データは現金に変更
                    data_list.append('1')
                else:
                    if row[12] == '01':
                        data_list.append('1')
                    else:
                        if row[12] == '02':
                            data_list.append('2')
                        else:
                            data_list.append('F')                            
                #景品のデータは現金に切替
                if row[12] == '00':
                    data_list.append(5)
                else:                    
                    if row[12] == '01': #現金のデータは明細種別を"5"にセット    
                        data_list.append(5)
                    else:    
                        data_list.append(int(row[14])) #明細種別番号               
                    
                #景品のデータは1000円に切替
                if row[12] == '00':    
                    data_list.append(1000) #決済金額
                    sum_price += 1000
                else:
                    data_list.append(int(row[17])) #決済金額
                    sum_price += int(row[17])
                    
                # データベースに存在していないかチェック    
                ck_count = self.resdb.db_wcheck1(data_list)      

                #対象データが無ければ書き込み用配列にappend
                if len(ck_count) == 0:
                    #検索日付・時間セット
                    res_list = self.resdb.date_set(kyear,kmonth,kday,khour,kminute,ksecond)
                    data_list.append(res_list[0])#日付（整数）
                    #data_list.append(res_list[1])#日付（ハイフン付文字列）
                    data_list.append(res_list[2])#時間（文字列）
                    data_list.append(res_list[3])#時間（文字列） 
                    #曜日・祝日セット
                    res_list = self.resdb.week_set(kyear,kmonth,kday)
                    data_list.append(res_list[0])#曜日コード
                    data_list.append(res_list[1]) #祝日フラグ 祝日なら'1' それ以外は'0'  
                    data_list.append(res_list[2])#祝日名（空白有）        
                    
                    output_list.append(data_list)           
                    out_count += 1
                    #debug 
                    if in_count % 100 == 0: #100件処理毎に表示
                        print('データ入力件数',in_count)
                else:
                    out_err += 1        
                
        db_updatedate = date_time[0:4] + '-' + date_time[4:6] + '-' + date_time[6:8]
                    
        if out_err > 0:
                print('入力不可件数：',out_err)
        if out_count == 0 and out_err == 0:
            edit_status = 9
        else:
            edit_status = 0
            #Debug
            print('出力件数',out_count)
            print('合計金額',sum_price)
            # DBへデータ出力
            data_num = self.resdb.data_insert(output_list)               
        
        del self.resdb
        #debug
        dt_now = datetime.datetime.now()
        print('売上データ出力処理終了(売上明細)：',dt_now)     
        #
        return edit_status,data_num,out_err,db_updatedate
    
    #####################################
    # TOAMASファイル編集
    #####################################
    def toamas_edit(self):
        
        out_count = 0
        in_count = 0
        out_err = 0
        sum_price  = 0
        output_list = []       
        
        #debug
        dt_now = datetime.datetime.now()
        print('売上データ出力処理開始(TOAMAS)：',dt_now)     
        #
        
        
        for row in self.row:
            data_list = []
            if row[2] != '現金' and row[3] != '未了（不明）' and row[3] != '未了（未書込）' : #現段階では現金データは対象外とする。未了は対象外
                in_count += 1
                date_time = row[0]
                #売上日を日付と時間に分け,更に'/'と':'で分けて数字にする
                datetime_list = date_time.split()
                datetime_date = datetime_list[0].split('-')
                datetime_time = datetime_list[1].split(':')       
                        
                data_list.append(int(datetime_date[0]))#決済年
                data_list.append(int(datetime_date[1]))#決済月
                data_list.append(int(datetime_date[2]))#決済日
                data_list.append(int(datetime_time[0]))#決済時
                data_list.append(int(datetime_time[1]))#決済分
                data_list.append(int(datetime_time[2]))#決済秒
                
                kyear = int(datetime_date[0])#決済年
                kmonth = int(datetime_date[1])#決済月
                kday = int(datetime_date[2])#決済日
                khour = int(datetime_time[0])#決済時
                kminute = int(datetime_time[1])#決済分
                ksecond = int(datetime_time[2])#決済秒    
                
                # kessai_no += 1
                # if kessai_no >= 100:
                #     kessai_no = 1
                # data_list.append(int(kessai_no))#決済番号
                data_list.append(int(row[15])) #決済番号
                
                #設置場所資産番号から設置場所番号を検索
                ret_rows = self.resdb.set_placecd(row[8])
                data_list.append(ret_rows[0][0]) 
                #placecd = ret_rows[0]        
                
                #明細区分番号
                if row[2] == '現金':
                    data_list.append('1') 
                else:
                    data_list.append('2') 
                
                #明細種別名称から明細種別番号を検索
                ret_rows = self.resdb.set_meisaisyubetu(row[2])
                data_list.append(ret_rows[0][0])             
                
                kingaku_str = ""
                max_len = len(row[1])
                kingaku = row[1]
                if max_len >= 4:
                    for i in range(0,max_len):
                        if kingaku[i] != ',':
                            kingaku_str = kingaku_str + str(kingaku[i])
                else:
                    kingaku_str = int(row[1])
                    
                kingaku_dec = int(kingaku_str)
                #kingaku = row[1]
                #kingaku.replace(',','') 
                data_list.append(kingaku_dec) #決済金額
                sum_price += kingaku_dec
                # データベースに存在していないかチェック
                ck_count = self.resdb.db_wcheck2(data_list)            

                #対象データが無ければ書き込み用配列にappend
                if len(ck_count) == 0:
                    #検索日付・時間セット
                    res_list = self.resdb.date_set(kyear,kmonth,kday,khour,kminute,ksecond)
                    data_list.append(res_list[0])#日付（整数）
                    #data_list.append(res_list[1])#日付（ハイフン付文字列）
                    data_list.append(res_list[2])#時間（文字列）
                    data_list.append(res_list[3])#時間（文字列） 
                    #曜日・祝日セット
                    res_list = self.resdb.week_set(kyear,kmonth,kday)
                    data_list.append(res_list[0])#曜日コード
                    data_list.append(res_list[1]) #祝日フラグ 祝日なら'1' それ以外は'0'  
                    data_list.append(res_list[2])#祝日名（空白有）        
                    
                    output_list.append(data_list)            
                    out_count += 1
                    #debug 
                    if in_count % 100 == 0:
                        print('データ入力件数',in_count) #100件処理毎に表示
                else:
                    out_err += 1        
                
        db_updatedate = datetime_date[0] + '-' + datetime_date[1] + '-' + datetime_date[2]
                    
        if out_err > 0:
                print('入力不可件数：',out_err)
                
        if out_count == 0 and out_err == 0:
            edit_status = 9
        else:
            edit_status = 0
            #Debug
            print('出力件数',out_count)
            print('合計金額',sum_price)
            # DBへデータ出力
            data_num = self.resdb.data_insert(output_list)           
        
        del self.resdb
        #debug
        dt_now = datetime.datetime.now()
        print('売上データ出力処理終了(TOAMAS)：',dt_now)     
        #
        return edit_status,data_num,db_updatedate  
    #####################################
    # ヤマトフィナンシャルファイル編集
    #####################################
    def yamato_edit(self,parm):
        kessai_no = 0 
        output_list = []
        out_count = 0
        in_count = 0
        out_err = 0   
        
        for row in self.row:
            #対象データ抽出
            res = self.resdb.data_choice(row)
        
            # データが1階分だったら    
            if res != 9:
                in_count += 1
                data_list = []
                date_time_int = int(row[2])
                #対象期間の判定
                sdate = parm[5]
                edate = parm[6]
                if date_time_int >= int(sdate):
                    if date_time_int <= int(edate):
                        date_time = str(date_time_int)
                        #date_time = row[3]
                        data_list.append(int(date_time[0:4]))#決済年
                        data_list.append(int(date_time[4:6]))#決済月
                        data_list.append(int(date_time[6:8]))#決済日
                        data_list.append(int(date_time[8:10]))#決済時
                        data_list.append(int(date_time[10:12]))#決済分
                        data_list.append(int(date_time[12:14]))#決済秒
                        
                        kyear = int(date_time[0:4])#決済年
                        kmonth = int(date_time[4:6])#決済月
                        kday = int(date_time[6:8])#決済日
                        khour = int(date_time[8:10])#決済時
                        kminute = int(date_time[10:12])#決済分
                        ksecond = int(date_time[12:14])#決済秒    
                        #data_list.append(int(date_time[14:16]))#決済番号
                        kessai_no += 1
                        data_list.append(kessai_no)#決済番号
                                            
                        data_list.append(1) #設置場所番号（1階限定）         
                    
                        data_list.append('2') #明細区分番号(この場合は固定)
                        
                        data_list.append(int(res))  #対象データ抽出のリターンを明細種別にセット         
                        data_list.append(int(row[3])) #決済金額
                        sum_price += int(row[3])
                        # データベースに存在していないかチェック    
                        ck_count = self.resdb.db_wcheck2(data_list)       

                        #対象データが無ければ書き込み用配列にappend
                        if len(ck_count) == 0:
                            #検索日付・時間セット
                            res_list = self.resdb.date_set(kyear,kmonth,kday,khour,kminute,ksecond)
                            data_list.append(res_list[0])#日付（整数）
                            #data_list.append(res_list[1])#日付（ハイフン付文字列）
                            data_list.append(res_list[2])#時間（文字列）
                            data_list.append(res_list[3])#時間（文字列） 
                            #曜日・祝日セット
                            res_list = self.resdb.week_set(kyear,kmonth,kday)
                            data_list.append(res_list[0])#曜日コード
                            data_list.append(res_list[1]) #祝日フラグ 祝日なら'1' それ以外は'0'  
                            data_list.append(res_list[2])#祝日名（空白有）        
                            
                            output_list.append(data_list)            
                            out_count += 1
                            #debug 
                            if in_count % 100 == 0:
                                print('データ入力件数',in_count) #100件処理毎に表示
                        else:
                            out_err += 1
                    else:
                        out_err += 1        
                
        db_updatedate = date_time[0:4] + '-' + date_time[4:6] + '-' + date_time[6:8]
                    
        if out_err > 0:
                print('入力不可件数：',out_err)
        
        if out_count == 0 and out_err == 0:
            edit_status = 9
        else:
            edit_status = 0
            # DBへデータ出力
                     #Debug
            print('出力件数',out_count)
            print('合計金額',sum_price)
            data_num = self.resdb.data_insert(output_list)   
        
        del self.resdb
        return edit_status,data_num,db_updatedate       
    
        
