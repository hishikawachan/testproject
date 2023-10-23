# -*- coding: utf-8 -*-
# ======================================
# 
# 電子マネー管理システム　フレームワーク
# Excel セル操作モジュール
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

import openpyxl

class excel_operate:
   def __init__(self, excel_dir):
      self.excel_dir = excel_dir
      self.workbook = openpyxl.load_workbook(self.excel_dir)
      self.sheet_number = len(self.workbook.get_sheet_names())
      self.sheet_contains =  [self.workbook.worksheets[i] for i in range(self.sheet_number)]
      self.merged_cells_list = [self.sheet_contains[i].merged_cells.ranges for i in range(self.sheet_number)]

   def get_merged_cells_location(self):
      self.merged_cells_location_list = [""]*self.sheet_number
      for i in range(self.sheet_number):
         self.merged_cells_location_list[i] = [format(self.merged_cells_list[i][j]) for j in range(len(self.merged_cells_list[i]))]
      else:
         pass
         #print("The merged cells location was got!")
         #print("***********************************")


   def break_merged_cells(self):
      for i in range(self.sheet_number):
         for j in range(len(self.merged_cells_list[i])):
            self.sheet_contains[i].unmerge_cells(self.merged_cells_location_list[i][j])
         else:
            pass
      else:
         #print('The all merged cells were unmerged!')
         pass
