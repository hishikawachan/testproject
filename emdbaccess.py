# -*- coding: utf-8 -*-
# ======================================
# 
# 電子マネー管理システム　フレームワーク
# DBアクセス
# 
# [環境]
#   Python 3.10.8
#   VSCode 1.64
#   <拡張>
#     |- Python  V2021.12
#     |- Pylance V2021.12
#
# [更新履歴]
#   2022/2/4  新規作成
# ======================================
#from urllib.parse import urlparse
import mysql.connector
#
# ＤＢアクセス管理クラス
#
class dbAccessor:

    # -----------------------------------
    # コンストラクタ
    #
    # コネクションを取得し、クラス変数にカーソルを保持する。
    # -----------------------------------
    def __init__(self, dbName, port, hostName, id, password):
        #print("start:__init__")
        
        try:
            # DBに接続する
            self.conn = mysql.connector.connect(
                host = hostName,
                port = port,
                user = id,
                password = password,
                database = dbName,
            )

            # コネクションの設定
            self.conn.autocommit = False
            
            #自動的に再接続する
            self.conn.ping(reconnect=True)

            # カーソル情報をクラス変数に格納
            self.conn.is_connected()
            self.cur = self.conn.cursor() 
            
            self.cur.execute("SHOW TABLES")
            self.table_name =[]
            for tt in self.cur:
                self.table_name.append(tt)   
           
        except (mysql.connector.errors.ProgrammingError) as e:
            print(e)

        #print("end:__init__")
    # -----------------------------------
    # テーブル一覧の取得
    #
    # クエリを実行し、テーブル一覧を取得する
    # -----------------------------------
    def table_name_get(self):
        
        return self.table_name

    # -----------------------------------
    # クエリの実行
    #
    # クエリを実行し、取得結果を呼び出し元に通知する。
    # -----------------------------------
    def excecuteQuery(self, sql):
        #print("start:excecuteQuery")

        try:
            self.cur.execute(sql)
            rows = self.cur.fetchall()
            return rows
        except (mysql.connector.errors.ProgrammingError) as e:
            print(e)

        #print("end:excecuteQuery")
    
    # -----------------------------------
    # インサートの実行
    #
    # インサートを実行する。
    # -----------------------------------
    def excecuteInsert(self, sql):
        #print("start:excecuteInsert")

        try:
            self.cur.execute(sql)
            self.conn.commit()
            return self.cur.rowcount
        except (mysql.connector.errors.ProgrammingError) as e:
            self.conn.rollback()
            print(e)

        #print("end:excecuteInsert")
    
    # -----------------------------------
    # 複数データ一括インサートの実行
    #
    # インサートを実行する。
    # -----------------------------------
    def excecuteInsertmany(self, sql, data):
        #print("start:excecuteInsertmany")

        try:
            self.cur.executemany(sql,data)
            self.conn.commit()
            return self.cur.rowcount
        except (mysql.connector.errors.ProgrammingError) as e:
            self.conn.rollback()
            print(e)

        #print("end:excecuteInsert")

    # -----------------------------------
    # アップデートの実行
    #
    # アップデートを実行する。
    # -----------------------------------
    def excecuteUpdate(self, sql):
        #print("start:excecuteUpdate")

        try:
            self.cur.execute(sql)
            self.conn.commit()
            return self.cur.rowcount
        except (mysql.connector.errors.ProgrammingError) as e:
            self.conn.rollback()
            print(e)

        #print("end:excecuteUpdate")

    # -----------------------------------
    # デリートの実行
    #
    # デリートを実行する。
    # -----------------------------------
    def excecuteDelete(self, sql):
        #print("start:excecuteDelete")

        try:
            self.cur.execute(sql)
            self.conn.commit()
            return self.cur.rowcount
        except (mysql.connector.errors.ProgrammingError) as e:
            self.conn.rollback()
            print(e)

        #print("end:excecuteDelete")

    # -----------------------------------
    # デストラクタ
    #
    # コネクションを解放する。
    # -----------------------------------
    def __del__(self):
        #print("start:__del__")
        try:
            self.conn.close()
        except (mysql.connector.errors.ProgrammingError) as e:
            print(e)
        #print("end:__del__")