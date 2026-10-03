# -*- coding: utf-8 -*-
"""
Created on Wed Aug  7 20:40:03 2024

@author: 陳皓均

託運估價系統視窗介面
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from freight_core import (ExcelFormatError, FreightTable, app_dir,
                          default_excel_path)

APP_TITLE = "託運估價系統 v4.0"
FONT = "微軟正黑體"

PRIMARY = "#1F5FAD"
MUTED = "#666666"
SUCCESS = "#1E7E34"
WARN = "#D26900"
ERROR = "#C0392B"
CARD_BG = "#F3F7FC"
CARD_BORDER = "#C9D8EC"

# 注意事項上色 {編號: 顏色}
NOTE_COLORS = {8: WARN, 9: WARN, 10: ERROR, 11: ERROR}

NORTH_TAB, SOUTH_TAB = 0, 1  # 分頁順序
MAX_KG_DIGITS = 6  # 重量最多輸入幾位數


# 介面

# 主視窗
class App:

    def __init__(self, root):
        self.root = root
        self.table = None
        self.quote = None

        root.title(APP_TITLE)
        root.resizable(False, False)
        self._setup_style()

        self.north_area = tk.IntVar(value=0)
        self.size = tk.IntVar(value=0)
        self.south_area = tk.IntVar(value=0)
        self.kg = tk.StringVar()
        for var in (self.north_area, self.size, self.south_area, self.kg):
            var.trace_add("write", lambda *_: self.refresh())

        self._build_header()
        body = ttk.Frame(root, padding=(16, 0, 16, 16))
        body.pack(fill="both", expand=True)
        body.rowconfigure(0, weight=1)
        self._build_tabs(body)
        self._build_side(body)

        self.load_table(default_excel_path(), show_error=False)

    def _setup_style(self):
        style = ttk.Style()
        style.configure(".", font=(FONT, 11))
        style.configure("TButton", padding=4)
        style.configure("TLabelframe.Label", font=(FONT, 11, "bold"))
        style.configure("TNotebook.Tab", font=(FONT, 11), padding=(16, 6))
        style.configure("Muted.TLabel", foreground=MUTED, font=(FONT, 9))

    # 版面

    # 標題列
    def _build_header(self):
        header = ttk.Frame(self.root, padding=(16, 14, 16, 10))
        header.pack(fill="x")

        titles = ttk.Frame(header)
        titles.pack(side="left")
        tk.Label(titles, text="託運估價系統", font=(FONT, 20, "bold"),
                 fg=PRIMARY).pack(anchor="w")
        ttk.Label(titles, text="選擇送貨地區與重量，馬上算出預估運費",
                  style="Muted.TLabel").pack(anchor="w")

        file_box = ttk.Frame(header)
        file_box.pack(side="right", anchor="s")
        ttk.Button(file_box, text="更換運費表…",
                   command=self.choose_file).pack(side="right")
        self.file_label = ttk.Label(file_box, text="")
        self.file_label.pack(side="right", padx=8)

    # 左側分頁
    def _build_tabs(self, parent):
        self.tabs = ttk.Notebook(parent)
        self.tabs.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        self.north_tab = ttk.Frame(self.tabs, padding=12)
        self.south_tab = ttk.Frame(self.tabs, padding=12)
        self.tabs.add(self.north_tab, text="北部（依板數）")
        self.tabs.add(self.south_tab, text="中南部（依重量）")
        self.tabs.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        tips = ttk.LabelFrame(parent, text=" 使用小提醒 ", padding=(10, 6))
        tips.grid(row=1, column=0, sticky="sew", padx=(0, 16), pady=(12, 0))
        for line in ("北部依板數查表，中南部依公斤數計算",
                     "選好條件，右側立即顯示運費，不用按計算",
                     "按「複製結果」可以直接貼到對帳表",
                     "多點送貨請依注意事項第 10、11 條自行加價"):
            ttk.Label(tips, text=f"• {line}").pack(anchor="w")

    # 右側結果與注意事項
    def _build_side(self, parent):
        side = ttk.Frame(parent)
        side.grid(row=0, column=1, rowspan=2, sticky="nsew")

        card = tk.Frame(side, bg=CARD_BG, highlightthickness=1,
                        highlightbackground=CARD_BORDER, padx=18, pady=14)
        card.pack(fill="x")
        tk.Label(card, text="預估運費", bg=CARD_BG, fg=MUTED,
                 font=(FONT, 11)).pack(anchor="w")
        self.price_label = tk.Label(card, text="—", bg=CARD_BG, fg=PRIMARY,
                                    font=(FONT, 30, "bold"))
        self.price_label.pack(anchor="w")
        self.detail_label = tk.Label(card, text="", bg=CARD_BG, justify="left",
                                     font=(FONT, 11))
        self.detail_label.pack(anchor="w")
        self.hint_label = tk.Label(card, text="", bg=CARD_BG, fg=WARN,
                                   justify="left", wraplength=380,
                                   font=(FONT, 10, "bold"))
        self.hint_label.pack(anchor="w", pady=(4, 0))
        self.copy_button = ttk.Button(card, text="複製結果",
                                      command=self.copy_result, state=tk.DISABLED)
        self.copy_button.pack(anchor="e")

        self.notes_box = ttk.LabelFrame(side, text=" 注意事項 ", padding=(10, 6))
        self.notes_box.pack(fill="both", expand=True, pady=(12, 0))

    # 單選按鈕群組
    def _radio_group(self, parent, title, options, variable, columns=1):
        box = ttk.LabelFrame(parent, text=f" {title} ", padding=(10, 6))
        box.pack(fill="x", pady=(0, 10))
        for i, text in enumerate(options):
            ttk.Radiobutton(box, text=text, value=i, variable=variable).grid(
                row=i // columns, column=i % columns, sticky="w",
                padx=(0, 18), pady=2)
        return box

    def _fill_north_tab(self):
        self._radio_group(self.north_tab, "送貨區域", self.table.north_areas,
                          self.north_area, columns=2)
        self._radio_group(self.north_tab, "重量／板數", self.table.sizes,
                          self.size, columns=2)

    def _fill_south_tab(self):
        self._radio_group(self.south_tab, "送貨區域", self.table.south_areas,
                          self.south_area, columns=2)

        box = ttk.LabelFrame(self.south_tab, text=" 重量 ", padding=(10, 8))
        box.pack(fill="x")
        row = ttk.Frame(box)
        row.pack(anchor="w")
        only_digits = (self.root.register(self._valid_kg), "%P")
        self.kg_entry = ttk.Entry(row, textvariable=self.kg, width=10,
                                  font=(FONT, 14), justify="right",
                                  validate="key", validatecommand=only_digits)
        self.kg_entry.pack(side="left")
        ttk.Label(row, text="KG").pack(side="left", padx=(6, 12))
        ttk.Button(row, text="清除", command=lambda: self.kg.set("")).pack(side="left")
        ttk.Label(box, style="Muted.TLabel", justify="left", text=(
            f"只能輸入整數公斤數\n"
            f"超過 {self.table.single_truck_limit:,} KG 自動以兩台車計算，"
            f"最多 {self.table.total_limit:,} KG")).pack(anchor="w", pady=(8, 0))

    def _fill_notes(self):
        for i, note in enumerate(self.table.notes, start=1):
            color = NOTE_COLORS.get(i, "#333333")
            ttk.Label(self.notes_box, text=note, foreground=color,
                      font=(FONT, 10), wraplength=380,
                      justify="left").pack(anchor="w", pady=1)

    # 運費表

    # 選擇運費表
    def choose_file(self):
        path = filedialog.askopenfilename(
            title="選擇運費表", initialdir=app_dir(),
            filetypes=[("Excel 檔案", "*.xlsx")])
        if path:
            self.load_table(path)

    # 讀取運費表並更新畫面
    def load_table(self, path, show_error=True):
        try:
            table = FreightTable(path)
        except FileNotFoundError:
            if show_error:
                messagebox.showerror("找不到檔案", f"找不到運費表：\n{path}")
            if self.table is None:
                self.file_label.config(text="找不到 運費.xlsx", foreground=ERROR)
                self._show_message("請按右上角「更換運費表…」選擇 Excel")
            return
        except ExcelFormatError:
            messagebox.showerror(
                "Excel 格式不符",
                "請確認 Excel 有「運費」與「注意事項」兩個工作表，"
                "且欄位位置沒有被搬動。")
            return

        self.table = table
        self.file_label.config(text=f"運費表：{table.name}", foreground=SUCCESS)
        for frame in (self.north_tab, self.south_tab, self.notes_box):
            for child in frame.winfo_children():
                child.destroy()
        self._fill_north_tab()
        self._fill_south_tab()
        self._fill_notes()
        for var in (self.north_area, self.size, self.south_area):
            var.set(0)
        self.refresh()

    # 計算與顯示

    # 重量只能輸入整數
    @staticmethod
    def _valid_kg(text):
        return text == "" or (text.isdigit() and len(text) <= MAX_KG_DIGITS)

    # 切換分頁
    def _on_tab_changed(self, event=None):
        if self.table and self.tabs.index("current") == SOUTH_TAB:
            self.kg_entry.focus_set()
        self.refresh()

    # 重新計算運費
    def refresh(self):
        if self.table is None:
            return
        if self.tabs.index("current") == NORTH_TAB:
            self._show_quote(self.table.north_quote(self.north_area.get(),
                                                    self.size.get()))
        elif self.kg.get():
            self._show_quote(self.table.south_quote(self.south_area.get(),
                                                    int(self.kg.get())))
        else:
            self._show_message("請輸入重量")

    def _show_quote(self, quote):
        self.quote = quote
        self.detail_label.config(
            text=f"區域：{quote.area}\n{quote.load_label}：{quote.load}")
        self.copy_button.config(state=tk.NORMAL)
        if quote.ok:
            self.price_label.config(text=f"{quote.price:,} 元", fg=PRIMARY)
            self.hint_label.config(
                text=(f"超過單車上限 {self.table.single_truck_limit:,} KG，"
                      "已以兩台車計算" if quote.trucks == 2 else ""), fg=WARN)
        else:
            self.price_label.config(text="無法計算", fg=ERROR)
            self.hint_label.config(text=quote.message, fg=ERROR)

    # 顯示提示文字
    def _show_message(self, text):
        self.quote = None
        self.price_label.config(text="—", fg=PRIMARY)
        self.detail_label.config(text=text)
        self.hint_label.config(text="")
        self.copy_button.config(state=tk.DISABLED)

    # 複製結果到剪貼簿
    def copy_result(self):
        if self.quote is None:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.quote.summary())
        self.copy_button.config(text="已複製 ✓")
        self.root.after(1500, lambda: self.copy_button.config(text="複製結果"))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
