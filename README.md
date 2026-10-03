# 託運估價系統

讀取運費 Excel 表，選擇送貨區域和重量（或板數）後，馬上算出預估運費，協助公司核對託運廠商的請款金額。

## 功能

- 視窗介面，點選區域和重量就能算出金額，不用翻 Excel
- **北部區域**：選擇送貨區域＋重量／板數，直接查表
- **中南部**：選擇「彰化(含)以北」或「雲林(含)以南」，輸入公斤數，依重量級距計算
- 超過單車最大載重時，自動以兩台車計算
- 右側顯示 Excel「注意事項」工作表的內容（加價規則、地區範圍等）
- 運費改在 Excel 調整，不用改程式

---

## 安裝

### 1. 事前準備

| 項目 | 說明 |
|---|---|
| **Python 3.9 以上** | 到 [python.org](https://www.python.org/downloads/) 下載，或安裝 [Anaconda](https://www.anaconda.com/download) |
| **運費 Excel 檔** | 專案已附 `運費.xlsx`，格式見下方「Excel 格式」 |

### 2. 下載專案

```
git clone https://github.com/tedchen05/freight-cost-estimator.git
cd freight-cost-estimator
```

### 3. 安裝套件

在專案資料夾開啟終端機，執行：

```
pip install -r requirements.txt
```

> 使用 Anaconda 的話，請在「Anaconda Prompt」中執行。

---

## 執行

### 方法一：終端機

在專案資料夾執行：

```
python 託運估價系統v3.0.py
```

### 方法二：VS Code 按 F5

1. 用 VS Code 開啟專案資料夾（第一次開啟時，右下角會建議安裝 Python 擴充功能，按「安裝」）
2. 按 `Ctrl + Shift + P`，輸入 `Python: Select Interpreter`，選擇**有安裝上述套件的 Python**
3. 按 **F5**

在「執行與偵錯」面板（`Ctrl + Shift + D`）上方可以切換：

| 選項 | 作用 |
|---|---|
| 執行託運估價系統 | 開啟估價視窗（預設） |
| 執行目前開啟的檔案 | 執行目前正在編輯的檔案 |

---

## 使用方式

1. 開啟程式後，輸入要讀取的 Excel 名稱（**不含副檔名**），例如 `運費`，按「確認」或 Enter
2. 進入「車資估算表」視窗：
   - **北部區域**：左邊選擇區域、中間選擇重量／板數，按中間的「**計算**」
   - **中南部**：左邊選擇「彰化(含)以北」或「雲林(含)以南」，輸入公斤數（只能輸入整數），按左邊的「**計算**」
3. 結果顯示在右下方，按「**清除**」後可以重新計算

> 一趟送兩個點以上時，程式不會自動加價，請依「注意事項」自行加算。

---

## Excel 格式

Excel 必須是 `.xlsx`，放在專案資料夾（執行程式的位置），包含兩個工作表：

| 工作表 | 說明 |
|---|---|
| 第一個工作表（運費） | 第一列是送貨區域，第一欄是重量與板數；右側另有「中南區載運量」與「彰化(含)以北／雲林(含)以南」的級距費率 |
| `注意事項` | 一欄名為「注意事項」，每列一條說明，會顯示在視窗右側 |

- 可以修改運費數字和注意事項內容，但**不能變更工作表名稱**，也不要搬動欄位位置
- 修改前建議先備份，改完請確認金額正確

---

## 常見問題

**Q：輸入名稱後顯示「找不到檔案」？**
確認 Excel 副檔名是 `.xlsx`、輸入時沒有加上 `.xlsx`，並且檔案和程式放在同一個資料夾。

**Q：出現 `ModuleNotFoundError: No module named 'pandas'`？**
執行用的 Python 沒有安裝套件。照「方法二」重新選擇 Python，或對目前的 Python 執行 `pip install -r requirements.txt`。

**Q：出現 `Missing optional dependency 'openpyxl'`？**
pandas 讀取 `.xlsx` 需要 openpyxl，執行 `pip install -r requirements.txt` 即可。

**Q：中南部計算顯示「請輸入數字」？**
重量只能輸入整數，不能有小數點、空白或單位。

**Q：顯示「重量太大無法計算」？**
重量超過兩台車的載重上限，請拆單或直接跟託運廠商議價。

---

## 檔案說明

```
freight-cost-estimator/
├── 託運估價系統v3.0.py    # 主程式（執行這個）
├── 操作手冊.txt           # 給使用者的簡易操作說明
├── requirements.txt       # 需要安裝的套件
├── 運費.xlsx              # 運費表（程式讀取的資料）
└── .vscode/
    ├── launch.json        # VS Code 按 F5 的執行設定
    └── extensions.json    # 建議安裝的 VS Code 擴充功能
```
