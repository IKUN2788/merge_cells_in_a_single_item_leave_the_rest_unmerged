# Python实战：打造高效Excel数据合并工具 (PyQt5 + Pandas)

在日常的数据处理工作中，我们经常遇到需要对Excel表格进行“合并单元格”操作的场景。例如，物流对账单、订单列表等，往往有多行数据属于同一个订单（如运单号相同），但具体的费用项（运费、提货费等）是分开列出的。为了生成清晰的对账单，我们需要将相同订单的信息合并显示，而保留明细行。

手动操作不仅费时费力，还容易出错。本文将介绍如何使用 Python (Pandas, Openpyxl) 和 PyQt5 开发一个带有图形界面的 Excel 合并工具，实现自动化处理。

## 1. 项目背景与需求

**痛点：**
- 原始数据通常是“一维”的清单，重复信息（如订单号、日期）在每一行都显示。
- 财务或业务部门需要查看“合并版”报表，即相同的订单信息合并单元格，右侧展示明细。
- 需要同时生成 JSON 数据供其他系统使用。

**解决方案：**
开发一个桌面小工具，用户只需选择 Excel 文件，勾选作为“合并依据”的列（Key），程序自动完成分组、合并和导出。

## 2. 核心功能

1.  **图形化界面 (GUI)**：基于 PyQt5，操作简单直观。
2.  **灵活的列选择**：自动读取 Excel 表头，用户可勾选哪些列作为 Key（合并依据），未勾选的列作为 Detail（明细）。
3.  **智能日期处理**：自动识别包含“日期”、“时间”、“Date”、“Time”的列，并统一格式化为 `YYYY-MM-DD`，解决 Excel 数字序列号（如 45932）的问题。
4.  **数据清洗**：自动处理科学计数法（如 `1.23E+11`），去除无效的空行。
5.  **双重输出**：
    - **Excel**：生成的表格中，Key 列自动合并单元格，且居中显示。
    - **JSON**：生成结构化的 JSON 数据，方便后续 API 调用或存档。

## 3. 技术栈

- **Python 3.x**
- **Pandas**: 强大的数据处理库，用于读取和分组数据。
- **Openpyxl**: 用于操作 Excel 文件，核心的 `merge_cells` 功能依赖它。
- **PyQt5**: 构建桌面应用程序界面。
- **Calamine**: (可选) 配合 Pandas 使用的高性能 Excel 读取引擎。

## 4. 实现细节

### 4.1 数据读取与预处理

使用 `pandas` 读取 Excel，为了防止精度丢失（如长数字变成科学计数法），我们统一读取为字符串格式，后续再按需转换。

```python
# 核心代码片段
self.df = pd.read_excel(
    self.file_path, 
    sheet_name=sheet_name, 
    header=header_row,
    engine='calamine', # 加速读取
    dtype=str          # 强制字符串，避免精度丢失
)
```

### 4.2 智能日期格式化

Excel 中的日期有时会存储为浮点数（序列号），有时是字符串。我们需要一个健壮的函数来统一它们。

```python
def format_excel_date(value: Any) -> str:
    """
    处理Excel日期格式，支持:
    1. 字符串 '2023-01-01'
    2. datetime 对象
    3. Excel序列号 (float/int) e.g. 45932.0
    """
    # ... 省略部分判断代码 ...
    
    # 针对 Excel 序列号的处理
    if isinstance(value, (int, float)):
        try:
            dt = datetime(1899, 12, 30) + timedelta(days=float(value))
            return dt.strftime('%Y-%m-%d')
        except Exception:
            return str(value)
            
    # ... 其他解析逻辑 ...
```

在最新的更新中，我们还加入了自动关键词检测：

```python
# 自动识别日期列
date_keywords = ['日期', '时间', 'Date', 'Time']
for col in processed_df.columns:
    if any(keyword in str(col) for keyword in date_keywords):
        processed_df[col] = processed_df[col].apply(format_excel_date)
```

### 4.3 分组与合并逻辑

这是工具的核心。我们需要根据用户选定的 Key 列，将数据分组。

1.  **构建 Key**：将每一行选中的 Key 列组合成一个 tuple。
2.  **分组**：遍历数据，将相同 Key 的行归类到同一个列表。
3.  **写入与合并**：使用 `openpyxl` 写入数据，并根据每组的起始行和结束行调用 `ws.merge_cells`。

```python
# 写入逻辑示意
for key_tuple, items in groups.items():
    start_row = row_num
    
    # 写入多行明细
    for item in items:
        # ... 写入 Detail 列 ...
        row_num += 1
        
    end_row = row_num - 1
    
    # 合并 Key 列
    if end_row > start_row:
        for i in range(len(key_columns)):
            ws.merge_cells(start_row=start_row, start_column=i+2, end_row=end_row, end_column=i+2)
```

### 4.4 JSON 输出

除了 Excel，工具还会生成如下结构的 JSON，Key 是合并列的组合字符串，Value 是明细列表：

```json
{
    "2023-10-21_SF123456_上海_北京": [
        {
            "序号": 1,
            "费用": 100.0,
            "服务": "运费"
        },
        {
            "序号": 2,
            "费用": 20.0,
            "服务": "包装费"
        }
    ]
}
```

## 5. 界面预览

(此处可插入界面截图)

界面主要分为三部分：
1.  **顶部**：选择文件、Sheet页及表头行数。
2.  **中部**：左侧勾选需要合并的列（Key），右侧预览数据。
3.  **底部**：一键生成结果。

## 6. 总结

通过 Python 编写这个小工具，我们极大地提高了对账单整理的效率。它展示了 Python 在办公自动化（RPA）领域的强大能力——不仅能处理数据，还能通过 GUI 方便非技术人员使用。

如果你也有类似的需求，不妨尝试自己动手实现一个！
