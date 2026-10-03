# -*- coding: utf-8 -*-
"""
託運估價系統
視窗介面：輸入 Excel 名稱後，選擇送貨區域與重量（或板數）估算託運運費。

@author: 陳皓均
"""

import tkinter as tk

from freight_core import FreightTable


# ===== 設定 =====

DEFAULT_EXCEL = "運費"

TITLE_FONT = ("TkDefaultFont", 16)
TEXT_FONT = ("TkDefaultFont", 12)
RESULT_FONT = ("TkDefaultFont", 18)

# 注意事項要上色的列：{列索引: 顏色}
NOTE_COLORS = {9: "#D26900", 10: "#D26900", 11: "#AE0000", 12: "#AE0000"}


# ===== 介面：讀取 Excel =====

def ask_excel():
    """輸入 Excel 名稱並讀取，成功回傳 FreightTable，關閉視窗則回傳 None"""
    loaded = {}

    root = tk.Tk()
    root.title("讀取 Excel 檔案")
    root.geometry("320x50")
    root.resizable(False, False)

    tk.Label(root, text="請輸入要讀取的Excel名稱:").grid(row=0, column=0)
    entry = tk.Entry(root, width=10)
    entry.insert(0, DEFAULT_EXCEL)
    entry.grid(row=0, column=1, padx=16)
    error = tk.Label(root, text="")
    error.grid(row=1, column=0, sticky="w")

    def on_submit(event=None):
        try:
            loaded["table"] = FreightTable(entry.get() + ".xlsx")
        except FileNotFoundError:
            error.config(text="找不到檔案")
            return
        except (ValueError, KeyError):
            error.config(text="Excel 格式不符")
            return
        root.destroy()

    tk.Button(root, text="確認", command=on_submit).grid(row=0, column=2)
    root.bind("<Return>", on_submit)
    entry.focus_set()

    root.mainloop()
    return loaded.get("table")


# ===== 介面：車資估算表 =====

class EstimatorWindow:
    """車資估算主視窗：左欄選區域與輸入重量、中欄選重量／板數、右欄顯示說明與結果"""

    def __init__(self, table):
        self.table = table

        self.win = tk.Tk()
        self.win.title("車資估算表")
        self.win.geometry("850x600")
        self.win.resizable(False, False)

        self.north_area = tk.IntVar(value=0)
        self.south_area = tk.IntVar(value=0)
        self.size = tk.IntVar(value=0)

        left = tk.Frame(self.win)
        middle = tk.Frame(self.win)
        right = tk.Frame(self.win)
        left.grid(row=0, column=0, sticky="nw")
        middle.grid(row=0, column=1, sticky="nw")
        right.grid(row=0, column=2, sticky="nw")

        self._build_left(left)
        self._build_middle(middle)
        self._build_right(right)

    def _title(self, parent, text):
        tk.Label(parent, text=text, font=TITLE_FONT).grid(sticky="w")

    def _radios(self, parent, options, variable):
        for i, text in enumerate(options):
            tk.Radiobutton(parent, text=text, value=i, variable=variable,
                           font=TEXT_FONT).grid(sticky="w")

    def _buttons(self, parent, on_calc, on_clear):
        """計算／清除按鈕列，回傳計算按鈕"""
        row = tk.Frame(parent)
        row.grid(sticky="w")
        calc = tk.Button(row, text="計算", command=on_calc)
        calc.pack(side="left")
        tk.Button(row, text="清除", command=on_clear).pack(side="left", padx=20)
        return calc

    def _build_left(self, frame):
        """北部區域、中南部區域與重量輸入"""
        self._title(frame, "請選擇區域")
        self._radios(frame, self.table.north_areas, self.north_area)

        self._title(frame, "")
        self._title(frame, "中南部")
        self._radios(frame, self.table.south_areas, self.south_area)

        self._title(frame, "")
        self._title(frame, "請輸入重量(KG)")
        self.kg_entry = tk.Entry(frame)
        self.kg_entry.grid(sticky="w")
        self.south_calc = self._buttons(frame, self.calc_south, self.clear_south)

    def _build_middle(self, frame):
        """北部用的重量／板數"""
        self._title(frame, "請選擇重量或板數")
        self._radios(frame, self.table.sizes, self.size)
        self.north_calc = self._buttons(frame, self.calc_north, self.clear_north)

    def _build_right(self, frame):
        """注意事項與估算結果"""
        self._title(frame, "其他說明")
        for i, note in enumerate(self.table.notes):
            tk.Label(frame, text=note, font=TEXT_FONT,
                     fg=NOTE_COLORS.get(i, "black")).grid(sticky="w")

        self._title(frame, "估算表：")
        self.result = tk.Label(frame, text="", font=RESULT_FONT)
        self.result.grid(sticky="w")

    def _show(self, text):
        """顯示結果，並鎖住計算按鈕直到按下清除"""
        self.result.config(text=text)
        self.north_calc.config(state=tk.DISABLED)
        self.south_calc.config(state=tk.DISABLED)

    def _clear(self):
        self.result.config(text="")
        self.north_calc.config(state=tk.NORMAL)
        self.south_calc.config(state=tk.NORMAL)

    def calc_north(self):
        area = self.north_area.get()
        size = self.size.get()
        price = self.table.north_price(area, size)
        self._show(f"\n區域：{self.table.north_areas[area]}\n\n"
                   f"重量/板數：{self.table.sizes[size]}\n\n"
                   f"預估金額：{price}元")

    def clear_north(self):
        self._clear()

    def calc_south(self):
        kg_text = self.kg_entry.get()
        if not kg_text.isdigit():
            self._show("請輸入數字")
            return

        kg = int(kg_text)
        area_name = self.table.south_areas[self.south_area.get()]
        price = self.table.south_price(self.south_area.get(), kg)

        if price is None:
            self._show("重量太大無法計算")
        elif kg >= self.table.max_kg:
            self._show(f"\n區域：{area_name}\n重量：{kg} KG\n\n"
                       f"超過最大{self.table.max_kg}重量\n以兩台車計算價錢\n\n"
                       f"預估金額：{price}元")
        else:
            self._show(f"\n區域：{area_name}\n\n重量：{kg} KG\n\n"
                       f"預估金額：{price}元")

    def clear_south(self):
        self._clear()
        self.kg_entry.delete(0, tk.END)

    def run(self):
        self.win.mainloop()


def main():
    table = ask_excel()
    if table is not None:
        EstimatorWindow(table).run()


if __name__ == "__main__":
    main()
