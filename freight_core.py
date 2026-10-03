# -*- coding: utf-8 -*-
"""
託運估價系統核心功能
讀取運費 Excel 表，並依送貨區域與重量（或板數）計算運費。

@author: 陳皓均
"""

import pandas as pd


# ===== 設定 =====

NOTES_SHEET = "注意事項"

# 運費表（第一個工作表）的欄位位置
SIZE_COL = 0                      # 重量／板數
NORTH_AREA_COLS = slice(1, 9)     # 北部各送貨區域
SOUTH_WEIGHT_COL = 9              # 中南區載運量（重量級距）
SOUTH_AREA_COLS = slice(10, 12)   # 彰化(含)以北、雲林(含)以南


# ===== 運費計算 =====

def single_truck_price(kg, weights, rates):
    """一台車的運費：前兩級是固定金額，之後依重量級距乘上每公斤費率"""
    for i in range(5, 1, -1):
        if kg >= weights[i]:
            return int(kg * rates[i])
    if kg >= weights[1]:
        return int(rates[1])
    return int(rates[0])


def south_price(kg, weights, rates):
    """中南部運費：超過一台車載重以兩台車計算，超過兩台車載重回傳 None"""
    max_kg = weights[6]
    if kg >= (max_kg - 1) * 2:
        return None
    if kg >= max_kg:
        first_truck = int((max_kg - 1) * rates[5])
        return first_truck + single_truck_price(kg - max_kg + 1, weights, rates)
    return single_truck_price(kg, weights, rates)


# ===== 運費表 =====

class FreightTable:
    """運費 Excel 的內容與查詢"""

    def __init__(self, path):
        df = pd.read_excel(path)
        notes = pd.read_excel(path, sheet_name=NOTES_SHEET)

        self.sizes = list(df.iloc[:, SIZE_COL])

        north = df.iloc[:, NORTH_AREA_COLS]
        self.north_areas = list(north.columns)
        self._north_prices = north

        south = df.iloc[:, SOUTH_AREA_COLS]
        self.south_areas = list(south.columns)
        self._south_rates = [list(south[col]) for col in south.columns]
        self.south_weights = list(df.iloc[:, SOUTH_WEIGHT_COL])
        self.max_kg = self.south_weights[6]

        self.notes = list(notes[NOTES_SHEET])

    def north_price(self, area, size):
        """北部：依區域與重量／板數查表"""
        return self._north_prices.iloc[size, area]

    def south_price(self, area, kg):
        """中南部：依區域與公斤數計算"""
        return south_price(kg, self.south_weights, self._south_rates[area])
