# PMBigTable 使用指南

## 簡介

PMBigTable 是一個用於項目管理的 Python 類別,它可以幫助您創建工作分解結構 (WBS),計算關鍵路徑,生成甘特圖,並輸出詳細的 Excel 報告。本指南將說明如何使用 PMBigTable 類別來管理您的項目。

## 安裝

使用 PMBigTable 之前,請確保您已經安裝了以下 Python 庫:

```
pip install openpyxl
```

## 基本用法

1. 導入必要的模組並創建 PMBigTable 實例:

```python
from datetime import datetime
from pm_big_table import PMBigTable

holidays = [datetime(2024, 2, 8).date(), datetime(2024, 2, 9).date(), datetime(2024, 2, 12).date(), datetime(2024, 2, 13).date(), datetime(2024, 2, 14).date()]
pm_tool = PMBigTable(working_days_per_week=5, project_start_date=datetime(2024, 1, 2), duration_unit='day', display_unit='day', holidays=holidays)
```

2. 準備工作分解結構 (WBS) 數據:

```python
wbs_data = [
    {'WBS': '1', 'Task': 'Task A', 'Duration': 2, 'Predecessors': '', 'Owner': 'Alex'},
    {'WBS': '2', 'Task': 'Task B', 'Duration': 3, 'Predecessors': 'Task A', 'Owner': 'Bob'},
    {'WBS': '3', 'Task': 'Task C', 'Duration': 5, 'Predecessors': 'Task A', 'Owner': 'Chris'},
    # ... 添加更多任務 ...
]
```

3. 載入 WBS 數據並生成報告:

```python
pm_tool.load_wbs(wbs_data)
pm_tool.print_schedule()
pm_tool.print_gantt_chart()
pm_tool.print_critical_path()
pm_tool.print_detailed_schedule()
pm_tool.generate_excel("project_schedule.xlsx")
```

## 詳細功能說明

### 初始化 PMBigTable

創建 PMBigTable 實例時,您可以設置以下參數:

- `working_days_per_week`: 每週工作天數 (默認為 5)
- `project_start_date`: 項目開始日期 (默認為 2024 年 1 月 1 日)
- `duration_unit`: 持續時間單位,可以是 'day' 或 'hour' (默認為 'day')
- `display_unit`: 顯示單位,可以是 'day' 或 'week' (默認為 'day')
- `holidays`: 假期列表 (默認為空列表)

### 載入 WBS 數據

使用 `load_wbs` 方法載入 WBS 數據。每個任務應包含以下信息:

- `WBS`: WBS 編號
- `Task`: 任務名稱
- `Duration`: 任務持續時間
- `Predecessors`: 前置任務 (如果有多個,用逗號分隔)
- `Owner`: 任務負責人

### 生成報告

- `print_schedule()`: 打印簡單的項目進度表
- `print_gantt_chart()`: 打印文本形式的甘特圖
- `print_critical_path()`: 打印項目的關鍵路徑
- `print_detailed_schedule()`: 打印詳細的項目進度表
- `generate_excel(filename)`: 生成詳細的 Excel 報告

## Excel 報告說明

`generate_excel` 方法會創建一個包含以下內容的 Excel 文件:

- 任務列表,包括 WBS 編號、任務名稱、負責人、計劃開始和結束日期
- 每個工作日的計劃工時和實際工時
- 項目進度和掙值分析 (EVM) 相關數據

### 重要提示

在首次打開 Excel 文件後,請使用「查找和替換」功能刪除文件中所有的 '@' 符號。這是為了確保 Excel 公式能夠正確計算。

## 示例

以下是一個完整的使用示例:

```python
from datetime import datetime
from pm_big_table import PMBigTable

# 定義假期
holidays = [datetime(2024, 2, 8).date(), datetime(2024, 2, 9).date(), datetime(2024, 2, 12).date(), datetime(2024, 2, 13).date(), datetime(2024, 2, 14).date()]

# 創建 PMBigTable 實例
pm_tool = PMBigTable(working_days_per_week=5, project_start_date=datetime(2024, 1, 2), duration_unit='day', display_unit='day', holidays=holidays)

# 準備 WBS 數據
wbs_data = [
    {'WBS': '1', 'Task': 'Task A', 'Duration': 2, 'Predecessors': '', 'Owner': 'Alex'},
    {'WBS': '2', 'Task': 'Task B', 'Duration': 3, 'Predecessors': 'Task A', 'Owner': 'Bob'},
    {'WBS': '3', 'Task': 'Task C', 'Duration': 5, 'Predecessors': 'Task A', 'Owner': 'Chris'},
    {'WBS': '4', 'Task': 'Task D', 'Duration': 2, 'Predecessors': 'Task A', 'Owner': 'Alex'},
    {'WBS': '5', 'Task': 'Task E', 'Duration': 3, 'Predecessors': 'Task D', 'Owner': 'Dell'},
    {'WBS': '6', 'Task': 'Task F', 'Duration': 4, 'Predecessors': 'Task D', 'Owner': 'Eric'},
    {'WBS': '7', 'Task': 'Task G', 'Duration': 1, 'Predecessors': 'Task B, Task F', 'Owner': 'Frank'},
    {'WBS': '8', 'Task': 'Task H', 'Duration': 20, 'Predecessors': 'Task G, Task E', 'Owner': 'Alex'},
    {'WBS': '9', 'Task': 'Task M', 'Duration': 6, 'Predecessors': 'Task H', 'Owner': 'Michael'},
]

# 載入 WBS 數據
pm_tool.load_wbs(wbs_data)

# 生成報告
pm_tool.print_schedule()
pm_tool.print_gantt_chart()
pm_tool.print_critical_path()
pm_tool.print_detailed_schedule()
pm_tool.generate_excel("project_schedule.xlsx")
```