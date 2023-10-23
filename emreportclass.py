# -*- coding: utf-8 -*-
# ======================================
# 電子マネー管理システム
# Mariaデータベースから帳票用データ加工
# Excelファイル出力モジュール呼び出し
#
# [環境]20
#   Python 3.10.8
#   VSCode 1.64
#   <拡張>
#     |- Python  V2021.12
#     |- Pylance V2021.12
#
# [更新履歴]
#   2023/2/19  新規作成
#   2023/9/12  出力フォルダー及びファイル名に会社コード追加
# ======================================
from datetime import datetime
import datetime
import os
import pandas as pd
from emdbclass import DataBaseClass
from emreportedit import dbReportEdit

######################################
# 初期化
#
# パラメタ
# [0]:dbip
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
# 対象データリスト
#####################################

class dbReport:
    def __init__(self,parm_data):
        self.parm_data = parm_data
    
    ######################################  
    # 設置場所別集計表出力
    ######################################   
    def place_print(self,df_paylog,flg):
        if flg == '1':
            sheet_name = '設置場所別(現金)'
        else:
            sheet_name = '設置場所別(電子決済)'
        dfw0 = df_paylog[df_paylog['paykbncd'] == flg]
        if len(dfw0) > 0: 
            df_sum_place = dfw0[['placename','payprice']].groupby('placename').sum()
            # 設置場所別集計表作成
            ret = self.res_ed.print_place(df_sum_place,sheet_name) 
            return ret 
    ######################################  
    # 金種別集計表出力
    ######################################   
    def kinsyu_print(self,df_paylog,flg):
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
            ret = self.res_ed.print_kinsyu(dfgp,dfgp2,sheet_name)         
            return ret
    
    ######################################  
    # 金種別集計表出力
    # 新バージョン 2023/9/22
    ######################################   
    def kinsyu_print2(self,data_kinsyu,kinsyu_sum_data,flg):
        if flg == '1':
            sheet_name = '金種別(現金)'
        else:
            sheet_name = '金種別(電子決済)'
        
        # 金種別集計表作成
        ret = self.res_ed.print_kinsyu2(data_kinsyu,kinsyu_sum_data,sheet_name)       
          
        return ret
    ######################################  
    # 時間別集計表出力
    ######################################   
    def jikan_print(self,df_paylog,flg): 
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
            ret = self.res_ed.print_jikan(dfgp,sheet_name)          
            return ret
    ######################################  
    # 設置場所別・時間別集計表出力（テストトライアル）
    ######################################   
    def jikan_print2(self,df_paylog,flg): 
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
            ret = self.res_ed.print_jikan2(dfgp,sheet_name)          
            return ret                    
    ######################################  
    # メイン
    ######################################           
    def main(self):  
            
        #データベース操作クラス初期化
        resdb = DataBaseClass(self.parm_data)
    
        input_symd = (self.parm_data[5].year * 10000)+(self.parm_data[5].month * 100)+(self.parm_data[5].day)
        input_eymd = (self.parm_data[6].year * 10000)+(self.parm_data[6].month * 100)+(self.parm_data[6].day)
        
        
        SYEAR = self.parm_data[5].year
        SMONTH = self.parm_data[5].month
        SDAY = self.parm_data[5].day
        EYEAR = self.parm_data[6].year
        EMONTH = self.parm_data[6].month
        EDAY = self.parm_data[6].day
        
        companycd = self.parm_data[8]
        
        #会社データ取得(指定)
        ret_rows = resdb.company_data_get(companycd)
        
        prec = ret_rows[0][2]
        block = ret_rows[0][3]
        
        dt_now = datetime.datetime.now()
        print('帳票出力処理開始：',dt_now) 
        
        #出力先パスの生成
        dir_date = str(companycd) + '_'+str(SYEAR)+str(SMONTH)+str(SDAY)+'_'+str(EYEAR)+str(EMONTH)+str(EDAY)
        dir_out_filepath = os.path.join(self.parm_data[9], companycd, dir_date)     
        # ディレクトリー存在チェック
        if os.path.exists(dir_out_filepath):
            pass
        else:
            os.mkdir(dir_out_filepath)       
        excel_file =  str(companycd) + '_'+str(SYEAR)+str(SMONTH)+str(SDAY)+'_'+str(EYEAR)+str(EMONTH)+str(EDAY)+'.xlsx'
        file_out_path = os.path.join(dir_out_filepath, excel_file)
    
        #気象データの更新
        res_list1 = resdb.weather_data_output(prec,block,SYEAR,SMONTH)
        #debug
        print(f'気象データ削除１件数：{res_list1[1]} 出力件数：{res_list1[0]}')
        #debug
        #年跨ぎ、月跨ぎの場合
        if (SYEAR != EYEAR) or (SMONTH != EMONTH) :
            res_list2 = resdb.weather_data_output(prec,block,EYEAR,EMONTH)
            #debug
            print(f'気象データ削除２件数：{res_list2[1]} 出力件数：{res_list2[0]}')
            #debug
        
        #帳票用元データの作成
        ret_syubetsu = resdb.syubetsu_get()
        ret_kbn = resdb.kbn_get()
        ret_place = resdb.place_get()
        ret_paylog = resdb.paylog_get(companycd,self.parm_data[5],self.parm_data[6])
        
        #金種別データ取得新バージョン2023/9/22
        # if self.test_flg == '1':
        #     #現金データ作成
        #     kbn = '1'
        #     ret_kinsyu1,kinsyu_sum_data1 = resdb.kinsyu_dataget(companycd,input_symd,input_eymd,kbn)
        #     #電子決済データ作成
        #     kbn = '2'
        #     ret_kinsyu2,kinsyu_sum_data2 = resdb.kinsyu_dataget(companycd,input_symd,input_eymd,kbn)
        
        # データ編集/出力クラス初期化
        del resdb
        self.res_ed = dbReportEdit(self.parm_data,file_out_path,SYEAR,SMONTH,SDAY,EYEAR,EMONTH,EDAY,prec,block)
        
        #決済種別別集計出力
        ret = self.res_ed.print_syubetsu(ret_syubetsu,ret_paylog)
        
        # 設置場所別データの生成・出力
        # 現金データ(100,500円,1000円)の抽出
        ret = self.place_print(ret_paylog,'1') 
        # 電子決済データの抽出  
        ret = self.place_print(ret_paylog,'2') 
        
        # 金種別データの生成・出力
        # 現金データの抽出
        ret = self.kinsyu_print(ret_paylog,'1')       
        # 電子決済データの抽出
        ret = self.kinsyu_print(ret_paylog,'2') 
        
        # 時間別データの生成
        # 現金データの抽出
        ret = self.jikan_print(ret_paylog,'1')  
        # 電子決済データの抽出
        ret = self.jikan_print(ret_paylog,'2') 
        
        # 出力したシートをPDFに変換
        self.res_ed.pdfconv(dir_out_filepath)                              
    
        dt_now = datetime.datetime.now()
        print('帳票処理終了：',dt_now)
        
    ###############################################################
    # ディストラクタ
    ###############################################################
    def __del__(self):
        #print('ディストラクタ呼び出し') 
        pass 
   
           
    