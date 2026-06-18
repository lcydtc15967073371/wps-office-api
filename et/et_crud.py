"""
WPS ET CRUD 操作示例 — Python COM 自动化
=====================================
要求: pip install pywin32
"""
import win32com.client, time

et = win32com.client.Dispatch("Ket.Application")
et.Visible = True
wb = et.Workbooks.Add()
ws = wb.Worksheets.Item(1)
ws.Name = "CRUD示例"

# ---- CREATE ----
headers = ["ID", "姓名", "部门", "薪资"]
for i, h in enumerate(headers):
    ws.Cells(1, i+1).Value = h
    ws.Cells(1, i+1).Font.Bold = True
    ws.Cells(1, i+1).Interior.Color = 0x4472C4
    ws.Cells(1, i+1).Font.Color = 0xFFFFFF

data = [(1, "张三", "技术部", 15000),
        (2, "李四", "市场部", 12000),
        (3, "王五", "财务部", 13000)]
for r, row in enumerate(data):
    for c, val in enumerate(row):
        ws.Cells(2+r, c+1).Value = val
time.sleep(0.5)
print("CREATE: 3 条记录已写入")

# ---- READ ----
ur = ws.UsedRange
print("READ: 当前数据")
for r in range(1, ur.Rows.Count+1):
    print("  " + " | ".join(str(ws.Cells(r,c).Text) for c in range(1, ur.Columns.Count+1)))

# ---- UPDATE ----
ws.Cells(3, 4).Value = 18000
print("UPDATE: 李四薪资 12000 -> 18000")
time.sleep(0.3)

# ---- DELETE ----
ws.Rows(4).Delete()
print("DELETE: 王五行已删除")
time.sleep(0.3)

# ---- READ AGAIN ----
print("FINAL: 最终数据")
ur = ws.UsedRange
for r in range(1, ur.Rows.Count+1):
    print("  " + " | ".join(str(ws.Cells(r,c).Text) for c in range(1, ur.Columns.Count+1)))
