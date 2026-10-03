# -*- coding: utf-8 -*-
"""
Created on Wed Aug  7 20:40:03 2024

@author: 陳皓均

託運估價系統核心功能
"""
import os
import re
import sys
from dataclasses import dataclass

import pandas as pd

DEFAULT_EXCEL = "運費.xlsx"
NOTES_SHEET = "注意事項"

# 運費表欄位位置
SIZE_COL = 0                      # 重量/板數
NORTH_AREA_COLS = slice(1, 9)     # 北部區域
SOUTH_WEIGHT_COL = 9              # 中南部重量級距
SOUTH_AREA_COLS = slice(10, 12)   # 中南部區域

# 注意事項編號，例如3.
NOTE_START = re.compile(r"^\d+\.")


# 檔案位置

# 程式所在資料夾(exe也適用)
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# 預設運費表路徑
def default_excel_path():
    return os.path.join(app_dir(), DEFAULT_EXCEL)


# 運費計算

# 一台車的運費
def single_truck_price(kg, weights, rates):
    for i in range(5, 1, -1):
        if kg >= weights[i]:
            return int(kg * rates[i])
    if kg >= weights[1]:
        return int(rates[1])
    return int(rates[0])


# 中南部運費，超過載重回傳None
def south_price(kg, weights, rates):
    max_kg = weights[6]
    if kg >= (max_kg - 1) * 2:
        return None
    if kg >= max_kg:
        first_truck = int((max_kg - 1) * rates[5])
        return first_truck + single_truck_price(kg - max_kg + 1, weights, rates)
    return single_truck_price(kg, weights, rates)


# 合併被換行拆開的注意事項
def merge_notes(lines):
    notes = []
    for line in lines:
        line = str(line).strip()
        if notes and not NOTE_START.match(line):
            notes[-1] += line
        else:
            notes.append(line)
    return notes


# 估價結果

# 無法計算時price為None
@dataclass
class Quote:
    area: str
    load_label: str     # 「重量/板數」或「重量」
    load: str
    price: int | None
    trucks: int = 1
    message: str = ""

    @property
    def ok(self):
        return self.price is not None

    # 複製用的一行文字
    def summary(self):
        if not self.ok:
            return f"{self.area}／{self.load}：{self.message}"
        trucks = "（兩台車）" if self.trucks == 2 else ""
        return f"{self.area}／{self.load}{trucks}：{self.price:,} 元"


# Excel格式不符
class ExcelFormatError(Exception):
    pass


# 運費表

# 讀取運費表並查詢運費
class FreightTable:

    def __init__(self, path):
        self.path = path
        try:
            df = pd.read_excel(path)
            notes = pd.read_excel(path, sheet_name=NOTES_SHEET)

            self.sizes = [str(s) for s in df.iloc[:, SIZE_COL]]

            north = df.iloc[:, NORTH_AREA_COLS]
            self.north_areas = [str(c) for c in north.columns]
            self._north_prices = north

            south = df.iloc[:, SOUTH_AREA_COLS]
            self.south_areas = [str(c) for c in south.columns]
            self._south_rates = [list(south[col]) for col in south.columns]
            self.south_weights = list(df.iloc[:, SOUTH_WEIGHT_COL])
            self.max_kg = int(self.south_weights[6])

            self.notes = merge_notes(notes[NOTES_SHEET].dropna())
        except (ValueError, KeyError, IndexError) as e:
            raise ExcelFormatError(str(e)) from e

    @property
    def name(self):
        return os.path.basename(self.path)

    # 單車載重上限
    @property
    def single_truck_limit(self):
        return self.max_kg - 1

    # 兩台車載重上限
    @property
    def total_limit(self):
        return (self.max_kg - 1) * 2 - 1

    # 北部：查表
    def north_quote(self, area, size):
        price = self._north_prices.iloc[size, area]
        quote = Quote(self.north_areas[area], "重量/板數", self.sizes[size], None)
        if pd.isna(price):
            quote.message = "運費表沒有這個組合的報價"
        else:
            quote.price = int(price)
        return quote

    # 中南部：依公斤數計算
    def south_quote(self, area, kg):
        price = south_price(kg, self.south_weights, self._south_rates[area])
        quote = Quote(self.south_areas[area], "重量", f"{kg:,} KG", price)
        if price is None:
            quote.message = f"超過兩台車上限 {self.total_limit:,} KG，請拆單或議價"
        elif kg > self.single_truck_limit:
            quote.trucks = 2
        return quote
