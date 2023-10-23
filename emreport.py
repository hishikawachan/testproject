# -*- coding: utf-8 -*-
# ======================================
# 電子マネー管理システム
# Mariaデータベースからデータ抽出
# 及びExcel及びPDF出力
# ※現金、電子決済別バージョン
# [環境]20
#   Python 3.10.8
#   VSCode 1.64
#   <拡張>
#     |- Python  V2021.12
#     |- Pylance V2021.12
#
# [更新履歴]
#   2023/2/19  新規作成
#   2023/9/12  出力DIR及びファイル名変更(会社コード追加)
#   2023/10/21 再構成
# ======================================
from datetime import datetime
import datetime
import os
import pandas as pd
import yaml
from emdbclass import DataBaseClass
from emreportedit import dbReportEdit

#################################################################
# DB制御classに渡すパラメータ
#################################################################
parm_data = []
#######################################
# 入力日付妥当性チェック
#######################################
def dateCheck(year,month,day):
    try:
        newDataStr="%04d/%02d/%02d"%(year,month,day)
        newDate=datetime.datetime.strptime(newDataStr,"%Y/%m/%d")
        return True
    except ValueError:
        return False
######################################
# 日付前後チェック
######################################
def dateNewCheck(syear,smonth,sday,eyear,emonth,eday):
    first = datetime.date(syear, smonth, sday)
    second = datetime.date(eyear, emonth, eday)
    seconds = (second - first).total_seconds()
    if seconds >= 0:        
        return True
    else:
        return False
######################################
# 処理対象拠点入力及びチェック
######################################
def company_check(dbed):
    #会社ID入力
    input_companyid = input('出力対象店舗ID(数字7桁)(9999999は終了):')
    if input_companyid != '9999999':        
        ret_rows = dbed.company_data_get(input_companyid)
        if len(ret_rows) > 0:
            print('出力対象店舗：',ret_rows[0][1])
            companycd = ret_rows[0][0]
            prec = ret_rows[0][2]
            block = ret_rows[0][3]
            #status = ret_rows[0]
            status =  5
        else:
            print('登録されている店舗ではありません')
            status =  3
    else:
        #print('処理を終了します')
        status = 9
    
    return status,companycd,prec,block
######################################
# 帳票出力対象日付入力及びチェック
######################################
def output_date_check():
    status = 0
    #対象日付入力 
    input_symd = input('出力対象の開始日付(yyyyMMdd 99999999は終了):')        
    input_eymd = input('出力対象の終了日付(yyyyMMdd 99999999は終了):')
    if (input_symd.isdecimal()  != True) or \
        (input_eymd.isdecimal()  != True):
        status = 3 #数字以外が入力
        print('数字以外が入力されました')
        return status,input_symd,input_eymd,status
    else:        
        if (int(input_symd) == 99999999) or \
            (int(input_eymd) == 99999999):
            status = 9 #終了処理
            print('終了が指示されました')
            return status,input_symd,input_eymd,status 
    
    SYEAR = int(input_symd[0:4])
    SMONTH = int(input_symd[4:6])
    SDAY = int(input_symd[6:8])
    EYEAR = int(input_eymd[0:4])
    EMONTH = int(input_eymd[4:6])
    EDAY = int(input_eymd[6:8])
     
     #日付の論理チェック   
    bool1 = dateCheck(SYEAR,SMONTH,SDAY)
    bool2 = dateCheck(EYEAR,EMONTH,EDAY)
    if (bool1 != True) or ( bool2 != True):
            print('正しい日付を入力してください') 
            status = 3
            return status,input_symd,input_eymd
    #日付の前後チェック
    bool = dateNewCheck(SYEAR, SMONTH, SDAY, EYEAR, EMONTH, EDAY)
    if bool == True:
        status = 5
        return status,input_symd,input_eymd
    else:
        print('日付の前後が間違っています')
        status = 3
        return status,input_symd,input_eymd
######################################  
# 設置場所別集計表出力
######################################   
def place_print(res_ed,df_paylog,flg):
    if flg == '1':
        sheet_name = '設置場所別(現金)'
    else:
        sheet_name = '設置場所別(電子決済)'
    dfw0 = df_paylog[df_paylog['paykbncd'] == flg]
    if len(dfw0) > 0: 
        df_sum_place = dfw0[['placename','payprice']].groupby('placename').sum()
        # 設置場所別集計表作成
        ret = res_ed.print_place(df_sum_place,sheet_name) 
        return ret 
######################################  
# 金種別集計表出力
######################################   
def kinsyu_print(res_ed,df_paylog,flg):
    if flg == '1':
        sheet_name = '金種別(現金)'
    else:
        sheet_name = '金種別(電子決済)'
    dfw0 = df_paylog[df_paylog['paykbncd'] == flg] 
    if len(dfw0) > 0:              
        dfw1 = dfw0[['paydatedec','payhour','payprice','paytimestr']]
        dfw2 = dfw1.astype({'paytimestr':str,'payprice': float,'payhour':int})  
        # 日付でソート
        dfw3 = dfw2.sort_values(by='paydatedec')
        # 時間でソート
        dfx = dfw3.sort_values(by='payhour') 
        #日付・時間で決済金額を集計                
        dfgp = pd.pivot_table(dfx, index=['paydatedec','payprice'], columns='payhour',aggfunc='count',margins=True,margins_name='Total')        
        #金種・日付で決済件数を集計
        dfgp2 = pd.pivot_table(dfx, index=['payprice'], columns='paydatedec',aggfunc='count',margins=True,margins_name='Total')
        # 金種別集計表作成
        ret = res_ed.print_kinsyu(dfgp,dfgp2,sheet_name)         
        return ret
######################################  
# 時間別集計表出力
######################################   
def jikan_print(res_ed,df_paylog,flg): 
    if flg == '1':
        sheet_name = '時間別(現金)'
    else:
        sheet_name = '時間別(電子決済)'
    dfw0 = df_paylog[df_paylog['paykbncd'] == flg] 
    if len(dfw0) > 0:
        dfw1 = dfw0[['paydatestr','paydatedec','payhour','payprice','paytimestr']] 
        dfw2 = dfw1.astype({'paydatedec':int,'paytimestr':str,'payprice': float,'payhour':int}) 
        #日付でソート 
        dfw3 = dfw2.sort_values(by='paydatedec')
        #時間でソート
        dfx = dfw3.sort_values(by='payhour')  
        dfgp = pd.pivot_table(dfx, index=['paydatedec'], columns='payhour',values=['payprice'],aggfunc='sum',margins=True) 
        # 時間別集計表作成
        ret = res_ed.print_jikan(dfgp,sheet_name)          
        return ret
######################################  
# 設置場所別・時間別集計表出力（テストトライアル）
######################################   
def jikan_print2(res_ed,df_paylog,flg): 
    if flg == '1':
        sheet_name = '時間別2(現金)'
    else:
        sheet_name = '時間別2(電子決済)'
    dfw0 = df_paylog[df_paylog['paykbncd'] == flg] 
    if len(dfw0) > 0:
        dfw1 = dfw0[['paydatestr','paydatedec','placename','payhour','payprice','paytimestr']] 
        dfw2 = dfw1.astype({'paydatedec':int,'paytimestr':str,'payprice': float,'payhour':int}) 
        #日付でソート 
        dfw3 = dfw2.sort_values(by='paydatedec')
        #時間でソート
        dfx = dfw3.sort_values(by='payhour')  
        dfgp = pd.pivot_table(dfx, index=['paydatedec','placename'], columns='payhour',values=['payprice'],aggfunc='sum',margins=True) 
        # 時間別集計表作成
        ret = res_ed.print_jikan2(dfgp,sheet_name)          
        return ret                    
######################################  
# メイン
######################################           
if __name__ == "__main__":  
    #　共通定義ファイル参照
    with open('C:/emoney/emoney.yaml','r',encoding="utf-8") as ry:
        config_yaml = yaml.safe_load(ry)
        parm_data.append(config_yaml['dbip'])
        parm_data.append(config_yaml['dbmarianame'])
        parm_data.append(config_yaml['dbport'])        
        parm_data.append(config_yaml['dbuser'])
        parm_data.append(config_yaml['dbpw'])
        parm_path = config_yaml['dir_filepath']#出力ファイル作成先パス
        
        #データベース操作クラス初期化
        resdb = DataBaseClass(parm_data)
    
        ret_row =  []
        ret_row.append(0)
        # 対象会社コード入力＋チェック(正常終了は5)
        while ret_row[0] < 5:
            ret_row = company_check(resdb)
        if ret_row[0] != 9: #終了ステータス＝終了以外なら            
            ret_row2 = []
            ret_row2.append(0)
            while ret_row2[0] < 5:
                #対象日付入力チェック(#正常終了は5)
                ret_row2 = output_date_check()
            if ret_row2[0] == 9:
                print('処理を終了します')
                del resdb              
        else:
            print('処理を終了します')
            del resdb
   
        if ret_row2[0] == 5: #初期処理正常なら
            input_symd = ret_row2[1]
            input_eymd = ret_row2[2]
            
            SYEAR = int(input_symd[0:4])
            SMONTH = int(input_symd[4:6])
            SDAY = int(input_symd[6:8])
            EYEAR = int(input_eymd[0:4])
            EMONTH = int(input_eymd[4:6])
            EDAY = int(input_eymd[6:8])
            
            companycd = ret_row[1]
            prec = ret_row[2]
            block = ret_row[3]
            
            dt_now = datetime.datetime.now()
            print('処理開始：',dt_now) 
            
            #出力先パスの生成
            dir_date = str(companycd)+'_'+str(SYEAR)+str(SMONTH)+str(SDAY)+'_'+str(EYEAR)+str(EMONTH)+str(EDAY)
            dir_out_filepath = os.path.join(parm_path, companycd, dir_date)     
            # ディレクトリー存在チェック
            if os.path.exists(dir_out_filepath):
                pass
            else:
                os.mkdir(dir_out_filepath)       
            excel_file = str(companycd)+'_'+str(SYEAR)+str(SMONTH)+str(SDAY)+'_'+str(EYEAR)+str(EMONTH)+str(EDAY)+'.xlsx'
            file_out_path = os.path.join(dir_out_filepath, excel_file)
        
            #気象データの更新
            res_list1 = resdb.weather_data_output(prec,block,SYEAR,SMONTH)
            #debug
            print(f'気象データ削除１件数：{res_list1[0]} 出力件数：{res_list1[1]}')
            #debug
            #年跨ぎ、月跨ぎの場合
            if (SYEAR != EYEAR) or (SMONTH != EMONTH) :
                res_list2 = resdb.weather_data_output(prec,block,EYEAR,EMONTH)
                #debug
                print(f'気象データ削除２件数：{res_list2[0]} 出力件数：{res_list2[1]}')
                #debug
            
            #帳票用元データの作成
            ret_syubetsu = resdb.syubetsu_get()
            ret_kbn = resdb.kbn_get()
            ret_place = resdb.place_get()
            ret_paylog = resdb.paylog_get(companycd,input_symd,input_eymd)
            
            
            # データ編集/出力クラス初期化
            del resdb
            res_ed = dbReportEdit(parm_data,file_out_path,SYEAR,SMONTH,SDAY,EYEAR,EMONTH,EDAY,prec,block)
            
            #決済種別別集計出力
            ret = res_ed.print_syubetsu(ret_syubetsu,ret_paylog)
            
            # 設置場所別データの生成・出力
            # 現金データ(100,500円,1000円)の抽出
            ret = place_print(res_ed,ret_paylog,'1') 
            # 電子決済データの抽出  
            ret = place_print(res_ed,ret_paylog,'2') 
            
            # 金種別データの生成・出力
            # 現金データの抽出
            ret = kinsyu_print(res_ed,ret_paylog,'1')       
            # 電子決済データの抽出
            ret = kinsyu_print(res_ed,ret_paylog,'2') 
            
            # 時間別データの生成
            # 現金データの抽出
            ret = jikan_print(res_ed,ret_paylog,'1')  
            # 電子決済データの抽出
            ret = jikan_print(res_ed,ret_paylog,'2') 
            # テストトライアル
            #ret = jikan_print2(res_ed,ret_paylog,'2')
            # テストトライアル  
            
            # 出力したシートをPDFに変換
            res_ed.pdfconv(dir_out_filepath)                              
        
        dt_now = datetime.datetime.now()
        print('処理終了：',dt_now)
   
           
    