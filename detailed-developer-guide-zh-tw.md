# 項目管理大表工具開發者指南

## 核心數據結構

### 1. wbs_data
- 類型：列表（List）of 字典（Dictionary）
- 用途：存儲工作分解結構（WBS）數據
- 結構：
  ```python
  [
    {
      'WBS': str,       # WBS 編號
      'Task': str,      # 任務名稱
      'Duration': int,  # 任務持續時間（天）
      'Predecessors': str,  # 前置任務，逗號分隔
      'Owner': str      # 任務負責人
    },
    # ...更多任務
  ]
  ```

### 2. task_schedule
- 類型：字典（Dictionary）
- 用途：存儲每個任務的詳細排程信息
- 結構：
  ```python
  {
    'Task Name': {
      'start': datetime,      # 任務開始時間
      'end': datetime,        # 任務結束時間
      'owner': str,           # 任務負責人
      'duration': int,        # 任務持續時間
      'ES': datetime,         # 最早開始時間
      'EF': datetime,         # 最早完成時間
      'LS': datetime,         # 最晚開始時間
      'LF': datetime,         # 最晚完成時間
      'Float': int            # 浮動時間
    },
    # ...更多任務
  }
  ```

### 3. owner_schedules
- 類型：字典（Dictionary）
- 用途：存儲每個負責人的工作時間表
- 結構：
  ```python
  {
    'Owner Name': {
      datetime: int,  # 日期: 工作小時數
      # ...更多日期
    },
    # ...更多負責人
  }
  ```

### 4. critical_path
- 類型：列表（List）
- 用途：存儲關鍵路徑上的任務名稱
- 結構：
  ```python
  ['Task A', 'Task B', 'Task C', ...]
  ```

## 主要類和方法

### PMBigTable 類

#### 初始化方法
```python
def __init__(self, working_days_per_week=5, project_start_date=datetime(2024, 1, 1), 
             duration_unit='day', display_unit='day', holidays=None):
    # ...
```
- `working_days_per_week`：每週工作天數
- `project_start_date`：項目開始日期
- `duration_unit`：持續時間單位（'day' 或 'hour'）
- `display_unit`：顯示單位（'day' 或 'week'）
- `holidays`：假期列表

#### 核心方法

1. `load_wbs(self, wbs_data)`
   - 載入 WBS 數據
   - 調用 `_calculate_schedule()` 和 `_calculate_critical_path()`

2. `_calculate_schedule(self)`
   - 計算項目排程
   - 填充 `task_schedule` 和 `owner_schedules`

3. `_calculate_critical_path(self)`
   - 計算關鍵路徑
   - 填充 `critical_path`

4. `print_schedule(self)`
   - 打印項目排程

5. `print_gantt_chart(self)`
   - 打印甘特圖

6. `print_critical_path(self)`
   - 打印關鍵路徑

7. `print_detailed_schedule(self)`
   - 打印詳細排程

8. `generate_excel(self, filename)`
   - 生成 Excel 報告

#### 輔助方法

1. `is_working_day(self, date)`
   - 檢查給定日期是否為工作日

2. `next_working_day(self, date)`
   - 獲取下一個工作日

3. `_find_next_available_date(self, start_date, owner)`
   - 找到下一個可用的開始日期

4. `_allocate_task(self, task_name, start_date, duration, owner)`
   - 分配任務到負責人的時間表

5. `_convert_duration(self, duration)`
   - 轉換持續時間單位

6. `_generate_timeline(self)`
   - 生成時間線字符串

7. `_date_range(self, start_date, end_date)`
   - 生成日期範圍

## 算法說明

### 1. 排程算法
- 按 WBS 順序遍歷任務
- 考慮前置任務和資源限制
- 使用 `_find_next_available_date()` 和 `_allocate_task()` 分配任務時間

### 2. 關鍵路徑算法
- 基於資源約束的關鍵路徑方法
- 所有任務按實際排程順序被視為關鍵路徑的一部分
- 計算每個任務的 ES、EF、LS、LF 和 Float

### 3. 甘特圖生成
- 使用字符繪圖方法
- '=' 表示工作日，'-' 表示非工作日

## 擴展建議

1. 多資源支持
   - 修改 `_calculate_schedule()` 以處理多個資源
   - 更新 `owner_schedules` 結構以支持多資源分配

2. 風險管理
   - 在 `wbs_data` 中添加風險屬性
   - 實現蒙特卡洛模擬方法

3. 成本管理
   - 添加成本相關屬性和計算方法
   - 擴展 EVM 計算

4. 用戶界面
   - 開發基於 Web 的界面（如使用 Flask 或 Django）
   - 實現交互式甘特圖編輯

5. 數據持久化
   - 實現項目數據的保存和加載功能
   - 考慮使用 SQLite 或其他數據庫

## 注意事項
- 維護現有 API 的兼容性
- 確保新功能不影響現有的性能
- 添加單元測試以確保代碼質量
- 更新文檔以反映新的變更和功能

