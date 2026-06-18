"""
WPS ET 批量导入 + 公式计算 — Python COM
"""
import win32com.client, time, csv

et = win32com.client.Dispatch("Ket.Application")
et.Visible = True

# 读取 CSV（示例数据）
wb = et.Workbooks.Add()
ws = wb.Worksheets.Item(1)

# 手动构建示例数据
rows_data = [
    ["产品", "Q1", "Q2", "Q3", "Q4"],
    ["A产品", 120, 150, 180, 200],
    ["B产品", 90, 110, 130, 160],
    ["C产品", 60, 80, 95, 120],
]

for r, row in enumerate(rows_data):
    for c, val in enumerate(row):
        ws.Cells(1+r, 1+c).Value = val

# 合计列
last_row = len(rows_data)
for r in range(2, last_row+1):
    ws.Cells(r, 6).Value = f"=SUM(B{r}:E{r})"
    ws.Cells(r, 6).Font.Bold = True

# 汇总行
ws.Cells(last_row+1, 1).Value = "合计"
ws.Cells(last_row+1, 1).Font.Bold = True
for c in range(2, 7):
    col = chr(64+c)
    ws.Cells(last_row+1, c).Value = f"=SUM({col}2:{col}{last_row})"
    ws.Cells(last_row+1, c).Font.Bold = True

# 格式化
ws.Range("A1:F1").Font.Bold = True
ws.Range("A1:F1").Interior.Color = 0x4472C4
ws.Range("A1:F1").Font.Color = 0xFFFFFF
ws.Columns.AutoFit()

print(f"批量导入完成: {last_row-1} 行数据 + 合计行")
print("合计列和汇总行均使用公式，自动计算")
