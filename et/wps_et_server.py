"""WPS ET 表格操作 MCP Server — COM 自动化（共享实例版）"""
import sys, os, re, tempfile, subprocess, time
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from mcp_base import MCPServer
import win32com.client, pythoncom

ET_PATH = r"E:\Program Files (x86)\Kingsoft\WPS Office\12.8.2.18205\office6\et.exe"

app = MCPServer("wps-et")

_et = None
_et_proc = None
_wb = None
_orig_path = None

def _ensure_et_running():
    global _et_proc, _et
    if _et is not None:
        try:
            _et.Visible
            return _et
        except Exception:
            _et = None
    _et_proc = subprocess.Popen([ET_PATH])
    for _ in range(30):
        try:
            _et = win32com.client.GetActiveObject("KET.Application")
            _et.Visible = True
            _et.DisplayAlerts = False
            return _et
        except Exception:
            time.sleep(1)
    raise RuntimeError("无法连接到 WPS ET (KET.Application)")

def _get_et():
    return _ensure_et_running()

def _get_ws():
    global _wb
    et = _get_et()
    if _wb is None:
        raise RuntimeError("没有打开的文件，请先调用 wps_open 打开一个 xlsx 文件")
    return _wb.ActiveSheet

def _col_to_num(col):
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - 64)
    return n

@app.tool("wps_open",
    "打开 xlsx 文件，共享实例供后续工具使用",
    input_schema={
        "type": "object",
        "properties": {
            "path": {"type": "string", "description": "xlsx 文件完整路径"},
            "sheet": {"type": "string", "description": "要激活的工作表名（可选，默认第一个）"}
        },
        "required": ["path"]
    })
def wps_open(args):
    global _wb, _orig_path
    # 关闭旧工作簿及所有残留临时文件
    try:
        et = _get_et()
        for wb in list(et.Workbooks):
            fn = wb.FullName
            if fn.endswith(".mimo.tmp"):
                wb.Close(SaveChanges=False)
                try: os.remove(fn)
                except: pass
    except Exception:
        pass
    _wb = None
    _orig_path = None
    path = os.path.abspath(args["path"])
    if not os.path.exists(path):
        raise RuntimeError(f"文件不存在: {path}")
    _orig_path = path
    basename = os.path.basename(path)
    tmp = os.path.join(tempfile.gettempdir(), f"~${basename}.mimo.tmp")
    import shutil
    for attempt in range(3):
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
            break
        except OSError:
            import time
            time.sleep(0.5)
    shutil.copy2(path, tmp)
    et = _get_et()
    _wb = et.Workbooks.Open(tmp)
    _wb.Saved = True
    sheet_name = args.get("sheet", "")
    if sheet_name:
        _wb.Sheets(sheet_name).Activate()
    ws = _wb.ActiveSheet
    sheets = [_wb.Sheets(j).Name for j in range(1, _wb.Sheets.Count+1)]
    return {"opened": path, "active_sheet": ws.Name, "sheets": sheets, "temp": tmp}

@app.tool("wps_new",
    "新建空白工作簿",
    input_schema={
        "type": "object",
        "properties": {"sheet_name": {"type": "string", "description": "工作表名（可选）"}}
    })
def wps_new(args):
    global _wb
    et = _get_et()
    _wb = et.Workbooks.Add()
    name = args.get("sheet_name", "")
    if name:
        _wb.ActiveSheet.Name = name
    return {"created": "空白工作簿", "sheet": _wb.ActiveSheet.Name}

def _next_sheet_name(name):
    n, i = name, 1
    while any(n == s.Name for s in _wb.Sheets):
        i += 1
        n = f"{name}({i})"
    return n

@app.tool("wps_add_sheet",
    "新建工作表",
    input_schema={
        "type": "object",
        "properties": {"name": {"type": "string", "description": "工作表名（可选）"}}
    })
def wps_add_sheet(args):
    ws = _wb.Sheets.Add()
    name = args.get("name", "")
    if name:
        ws.Name = _next_sheet_name(name)
    return {"created": ws.Name}

@app.tool("wps_copy_sheet",
    "复制工作表",
    input_schema={
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "要复制的工作表名（默认当前活动表）"},
            "new_name": {"type": "string", "description": "新表名（可选）"}
        }
    })
def wps_copy_sheet(args):
    name = args.get("name", _wb.ActiveSheet.Name)
    src = None
    for j in range(1, _wb.Sheets.Count+1):
        if str(_wb.Sheets.Item(j).Name) == name:
            src = _wb.Sheets.Item(j)
            break
    if src is None:
        raise RuntimeError(f"找不到工作表: {name}")
    src.Copy(None, _wb.Sheets(_wb.Sheets.Count))
    new_ws = _wb.ActiveSheet
    new_name = args.get("new_name", "")
    if new_name:
        new_ws.Name = _next_sheet_name(new_name)
    return {"copied": name, "as": new_ws.Name}

@app.tool("wps_extract_sheet",
    "复制指定区域内容到新工作表",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:J50"},
            "from_sheet": {"type": "string", "description": "源工作表（默认当前表）"},
            "new_sheet": {"type": "string", "description": "目标工作表名（可选，不传则新建）"}
        },
        "required": ["range"]
    })
def wps_extract_sheet(args):
    rng = args["range"]
    from_name = args.get("from_sheet", _wb.ActiveSheet.Name)
    ws_src = None
    for j in range(1, _wb.Sheets.Count+1):
        if str(_wb.Sheets.Item(j).Name) == from_name:
            ws_src = _wb.Sheets.Item(j)
            break
    if ws_src is None:
        raise RuntimeError(f"找不到工作表: {from_name}")
    cells = ws_src.Range(rng)
    cells.Copy()
    new_name = args.get("new_sheet", "")
    if new_name:
        ws_dst = None
        for j in range(1, _wb.Sheets.Count+1):
            if str(_wb.Sheets.Item(j).Name) == new_name:
                ws_dst = _wb.Sheets.Item(j)
                break
        if ws_dst is None:
            ws_dst = _wb.Sheets.Add()
            ws_dst.Name = _next_sheet_name(new_name)
        ws_dst.Activate()
    else:
        ws_dst = _wb.Sheets.Add()
    ws_dst.Range("A1").Select()
    ws_dst.Paste()
    return {"extracted": rng, "from": from_name, "to": ws_dst.Name}

@app.tool("wps_delete_sheet",
    "删除指定工作表",
    input_schema={
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "要删除的工作表名（默认当前表）"}
        }
    })
def wps_delete_sheet(args):
    name = args.get("name", _wb.ActiveSheet.Name)
    if _wb.Sheets.Count <= 1:
        raise RuntimeError("至少保留一个工作表")
    for j in range(1, _wb.Sheets.Count+1):
        s = _wb.Sheets.Item(j)
        if str(s.Name) == name:
            s.Delete()
            return {"deleted": name}
    raise RuntimeError(f"找不到工作表: {name}")

@app.tool("wps_list_sheets",
    "列出当前工作簿的所有工作表名")
def wps_list_sheets(args):
    global _wb
    if _wb is None:
        raise RuntimeError("没有打开的文件")
    sheets = [_wb.Sheets(j).Name for j in range(1, _wb.Sheets.Count+1)]
    return {"sheets": sheets, "count": len(sheets)}

@app.tool("wps_switch_sheet",
    "切换到指定工作表",
    input_schema={
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "工作表名"}
        },
        "required": ["name"]
    })
def wps_switch_sheet(args):
    global _wb
    if _wb is None:
        raise RuntimeError("没有打开的文件")
    name = args["name"]
    for j in range(1, _wb.Sheets.Count+1):
        s = _wb.Sheets.Item(j)
        if str(s.Name) == name:
            s.Activate()
            return {"active": str(s.Name)}
    raise RuntimeError(f"找不到工作表: {name}")

def _str(v):
    if v is None: return ""
    if isinstance(v, float) and v == int(v): return str(int(v))
    return str(v).strip()

@app.tool("wps_read",
    "读取区域数据（不传 range 默认为 A1开始的20行）",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:D10"}
        }
    })
def wps_read(args):
    rng = args.get("range", "A1:T20")
    ws = _get_ws()
    try:
        raw = _walk_com(ws, f"Range('{rng}').Value")
        if raw is None:
            cells = ws.Range(rng)
            r, c = cells.Rows.Count, cells.Columns.Count
            d = [[""] * c for _ in range(r)]
        elif isinstance(raw, tuple):
            d = [[_str(v) for v in row] for row in raw]
        else:
            d = [[_str(raw)]]
    except Exception:
        cells = ws.Range(rng)
        r, c = cells.Rows.Count, cells.Columns.Count
        d = [[_str(cells.Cells(i, j).Value) for j in range(1, c+1)] for i in range(1, r+1)]
    return {"r": len(d), "c": len(d[0]) if d else 0, "d": d}

@app.tool("wps_read_all",
    "读取当前工作表全部数据（token 较大）",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "可选区域如 A1:Z100，留空读全部"}
        }
    })
def wps_read_all(args):
    rng = args.get("range", "")
    ws = _get_ws()
    if not rng:
        cells = ws.UsedRange
        r, c = cells.Rows.Count, cells.Columns.Count
        rng = f"A1:{chr(64+c)}{r}"
    try:
        raw = _walk_com(ws, f"Range('{rng}').Value")
        if raw is None:
            cells = ws.Range(rng)
            r, c = cells.Rows.Count, cells.Columns.Count
            d = [[""] * c for _ in range(r)]
        elif isinstance(raw, tuple):
            d = [[_str(v) for v in row] for row in raw]
        else:
            d = [[_str(raw)]]
    except Exception:
        cells = ws.Range(rng)
        r, c = cells.Rows.Count, cells.Columns.Count
        d = [[_str(cells.Cells(i, j).Value) for j in range(1, c+1)] for i in range(1, r+1)]
    return {"r": len(d), "c": len(d[0]) if d else 0, "d": d}

@app.tool("wps_write",
    "写入二维数组到工作表",
    input_schema={
        "type": "object",
        "properties": {
            "start_cell": {"type": "string", "description": "起始单元格如 A1（默认 A1）"},
            "data": {"type": "array", "description": "二维数组", "items": {"type": "array", "items": {"type": "string"}}}
        },
        "required": ["data"]
    })
def wps_write(args):
    start = args.get("start_cell", "A1")
    data = args["data"]
    ws = _get_ws()
    m = re.match(r"([A-Z]+)(\d+)", start.upper())
    if not m: return {"error": f"无效单元格: {start}"}
    cs, rs = _col_to_num(m.group(1)), int(m.group(2))
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            ws.Cells(rs + r, cs + c).Value = val
    return {"written": f"{len(data)} rows x {max(len(r) for r in data)} cols"}

@app.tool("wps_update",
    "修改指定单元格的值",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "单元格如 B2"},
            "value": {"type": "string", "description": "新值"}
        },
        "required": ["cell", "value"]
    })
def wps_update(args):
    ws = _get_ws()
    ws.Range(args["cell"]).Value = args["value"]
    return {"updated": f"{args['cell']} = {args['value']}"}

@app.tool("wps_delete_row",
    "删除指定行",
    input_schema={
        "type": "object",
        "properties": {"row": {"type": "integer", "description": "行号从1开始"}},
        "required": ["row"]
    })
def wps_delete_row(args):
    row = args["row"]
    if not isinstance(row, int) or row < 1:
        raise ValueError(f"行号必须为正整数: {row}")
    _get_ws().Rows(row).Delete()
    return {"deleted": f"第{row}行"}

@app.tool("wps_save",
    "将临时副本的修改写回到原文件")
def wps_save(args):
    global _wb, _orig_path, _et
    if _wb is None:
        raise RuntimeError("没有打开的文件")
    if _orig_path is None:
        raise RuntimeError("没有原始文件路径，无法保存")
    tmp = _wb.FullName
    _wb.Save()
    try:
        import shutil
        shutil.copy2(tmp, _orig_path)
    except OSError:
        _wb.SaveAs(_orig_path)
        _wb.Close(SaveChanges=False)
        _wb = _et.Workbooks.Open(tmp)
        _wb.Saved = True
    return {"saved": _orig_path}

@app.tool("wps_format",
    "单元格格式：合并/取消/行高/列宽/字体/颜色/高亮重复/数据条/色阶/自动筛选",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:C10"},
            "merge": {"type": "boolean", "description": "合并单元格"},
            "unmerge": {"type": "boolean", "description": "取消合并"},
            "row_height": {"type": "number", "description": "行高（磅）"},
            "col_width": {"type": "number", "description": "列宽"},
            "font_bold": {"type": "boolean", "description": "加粗"},
            "font_size": {"type": "number", "description": "字号"},
            "font_color": {"type": "string", "description": "字体色 RGB 如 255,0,0"},
            "fill_color": {"type": "string", "description": "背景色 RGB 如 255,255,0"},
            "highlight_duplicates": {"type": "boolean", "description": "高亮重复项（黄色背景）"},
            "data_bar": {"type": "boolean", "description": "数据条"},
            "color_scale": {"type": "boolean", "description": "色阶"},
            "auto_filter": {"type": "boolean", "description": "启用自动筛选"},
        }
    })
def wps_format(args):
    ws = _get_ws()
    rng = args.get("range", "")
    obj = ws.Range(rng) if rng else ws.UsedRange
    if args.get("merge"): obj.Merge()
    if args.get("unmerge"): obj.UnMerge()
    if "row_height" in args: obj.RowHeight = args["row_height"]
    if "col_width" in args: obj.ColumnWidth = args["col_width"]
    if "font_bold" in args: obj.Font.Bold = args["font_bold"]
    if "font_size" in args: obj.Font.Size = args["font_size"]
    if "font_color" in args:
        rgb = [int(x.strip()) for x in args["font_color"].split(",")]
        obj.Font.Color = (rgb[2]<<16)+(rgb[1]<<8)+rgb[0]
    if "fill_color" in args:
        rgb = [int(x.strip()) for x in args["fill_color"].split(",")]
        obj.Interior.Color = (rgb[2]<<16)+(rgb[1]<<8)+rgb[0]
    if args.get("highlight_duplicates"):
        fc = obj.FormatConditions.AddUniqueValues()
        fc.DupeUnique = 1
        fc.Interior.Color = 65535
    if args.get("data_bar"):
        obj.FormatConditions.AddDatabar()
    if args.get("color_scale"):
        obj.FormatConditions.AddColorScale(ColorScaleType=3)
    if args.get("auto_filter"):
        ws.UsedRange.AutoFilter()
    elif "auto_filter" in args:
        ws.AutoFilterMode = False
    return {"formatted": rng or "UsedRange"}

@app.tool("wps_text_to_columns",
    "分列：将一列文本按分隔符拆分到多列",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "要拆分的列区域如 A1:A10"},
            "delimiter": {"type": "string", "description": "分隔符（默认逗号）"},
            "tab": {"type": "boolean", "description": "按制表符拆分"},
            "semicolon": {"type": "boolean", "description": "按分号拆分"},
        }
    })
def wps_text_to_columns(args):
    ws = _get_ws()
    rng = args.get("range", "")
    obj = ws.Range(rng) if rng else ws.UsedRange
    tab = args.get("tab", False)
    semicolon = args.get("semicolon", False)
    delim = args.get("delimiter", "")
    if tab:
        ws.Range(rng).TextToColumns(Destination=obj, DataType=1, TextQualifier=2, Tab=True)
    elif semicolon:
        ws.Range(rng).TextToColumns(Destination=obj, DataType=1, TextQualifier=2, Semicolon=True)
    elif delim:
        ws.Range(rng).TextToColumns(Destination=obj, DataType=1, TextQualifier=2, Other=True, OtherChar=delim)
    else:
        ws.Range(rng).TextToColumns(Destination=obj, DataType=1, TextQualifier=2, Comma=True)
    return {"text_to_columns": rng}

@app.tool("wps_sort",
    "排序（支持最多3列排序：key_column + key2 + key3）",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:C50"},
            "key_column": {"type": "string", "description": "排序列如 B"},
            "desc": {"type": "boolean", "description": "降序"},
            "key2": {"type": "string", "description": "第二排序列如 C"},
            "order2": {"type": "boolean", "description": "第二列降序"},
            "key3": {"type": "string", "description": "第三排序列如 D"},
            "order3": {"type": "boolean", "description": "第三列降序"},
            "has_header": {"type": "boolean", "description": "首行是标题行（默认自动检测）"}
        }
    })
def wps_sort(args):
    ws = _get_ws()
    rng = args.get("range", "")
    obj = ws.Range(rng) if rng else ws.UsedRange
    header = args.get("has_header")
    if header is None:
        first_cell = ws.Range(rng).Cells(1, 1).Value if rng else ws.UsedRange.Cells(1, 1).Value
        header = bool(first_cell and isinstance(first_cell, str) and not first_cell.lstrip("-").replace(".","").isdigit())
    h = 1 if header else 2
    row_num = 1
    m = re.match(r"^[A-Z]+(\d+)", rng)
    if m: row_num = int(m.group(1))
    def kr(k): return ws.Range(f"{k}{row_num}")
    key = args.get("key_column", "A")
    k2 = args.get("key2", "")
    k3 = args.get("key3", "")
    kwargs = {"Key1": kr(key), "Order1": 2 if args.get("desc") else 1, "Header": h}
    if k2:
        kwargs["Key2"] = kr(k2)
        kwargs["Order2"] = 2 if args.get("order2") else 1
    if k3:
        kwargs["Key3"] = kr(k3)
        kwargs["Order3"] = 2 if args.get("order3") else 1
    obj.Sort(**kwargs)
    return {"sorted": rng, "has_header": header}

@app.tool("wps_formula",
    "写入公式到单元格",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "单元格如 C1"},
            "formula": {"type": "string", "description": "公式如 =SUM(A1:A10)"}
        },
        "required": ["cell", "formula"]
    })
def wps_formula(args):
    ws = _get_ws()
    ws.Range(args["cell"]).Formula = args["formula"]
    return {"formula": args["cell"], "=": args["formula"]}

@app.tool("wps_count",
    "CountIf 统计 — 注意 * 通配符只匹配文本，不匹配数字单元格",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 H1:H110"},
            "criteria": {"type": "string", "description": "条件（如 \"绿色\" / \">100\" / \"*text*\"）"}
        },
        "required": ["range", "criteria"]
    })
def wps_count(args):
    ws = _get_ws()
    rng = ws.Range(args["range"])
    criteria = args["criteria"]
    result = _get_et().WorksheetFunction.CountIf(rng, criteria)
    return {"range": args["range"], "criteria": criteria, "count": result}

@app.tool("wps_insert",
    "插入行或列",
    input_schema={
        "type": "object",
        "properties": {
            "row": {"type": "integer", "description": "在第几行上方插入"},
            "col": {"type": "string", "description": "列如 C，在该列左侧插入"},
            "rows": {"type": "integer", "description": "插入 N 行（默认 1）"},
            "shift_down": {"type": "boolean", "description": "插入后下方单元格下移"}
        }
    })
def wps_insert(args):
    ws = _get_ws()
    if "row" in args:
        n = args.get("rows", 1)
        r = args["row"]
        ws.Rows(f"{r}:{r + n - 1}").Insert()
        return {"inserted": f"第{r}行前插入{n}行"}
    if "col" in args:
        c = _col_to_num(args["col"].upper())
        ws.Columns(c).Insert()
        return {"inserted": f"{args['col']}列前插入1列"}
    if args.get("shift_down"):
        ws.Range("A1").Insert(Shift=-4121)
        return {"inserted": "shift down"}
    return {"error": "请指定 row 或 col"}

@app.tool("wps_border",
    "单元格边框 — 支持 all/left/right/top/bottom/inner/outer + color(0,0,0) + weight(1-4) + style(1实线/-4115虚线)",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:C10"},
            "all": {"type": "boolean", "description": "全边框"},
            "left": {"type": "boolean", "description": "左边框"},
            "right": {"type": "boolean", "description": "右边框"},
            "top": {"type": "boolean", "description": "上边框"},
            "bottom": {"type": "boolean", "description": "下边框"},
            "inner": {"type": "boolean", "description": "内部边框"},
            "outer": {"type": "boolean", "description": "外边框"},
            "color": {"type": "string", "description": "RGB 如 0,0,0（默认黑色）"},
            "weight": {"type": "integer", "description": "线粗 1-4（默认 2）"},
            "style": {"type": "integer", "description": "线型 1=实线 -4115=虚线 -4119=双线 4=点划线 -4142=无线条"}
        }
    })
def wps_border(args):
    ws = _get_ws()
    rng = args.get("range", "")
    obj = ws.Range(rng) if rng else ws.UsedRange
    rgb = [int(x.strip()) for x in (args.get("color", "0,0,0").split(","))]
    color = (rgb[2]<<16)+(rgb[1]<<8)+rgb[0]
    weight = args.get("weight", 2)
    style = args.get("style", 1)
    sides = []
    if args.get("all"): sides = [7,8,9,10,11,12]  # xlEdgeLeft..xlInsideHorizontal
    if args.get("left"): sides.append(7)
    if args.get("right"): sides.append(10)
    if args.get("top"): sides.append(8)
    if args.get("bottom"): sides.append(9)
    if args.get("inner"): sides.extend([11,12])
    if args.get("outer"): sides.extend([7,8,9,10])
    for s in sides:
        obj.Borders(s).Color = color
        obj.Borders(s).Weight = weight
        obj.Borders(s).LineStyle = style
    return {"bordered": rng, "sides": sides}

@app.tool("wps_find_replace",
    "查找替换 — match_partial=true 支持部分匹配",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "查找范围（默认整表）"},
            "find": {"type": "string", "description": "要查找的内容"},
            "replace": {"type": "string", "description": "替换为（不传则只查找）"},
            "look_in": {"type": "integer", "description": "-4123=公式 -4163=值 -4144=批注（默认 -4163）"},
            "match_case": {"type": "boolean", "description": "区分大小写"},
            "match_partial": {"type": "boolean", "description": "部分匹配（默认 false=全单元格匹配）"}
        },
        "required": ["find"]
    })
def wps_find_replace(args):
    ws = _get_ws()
    rng = args.get("range", "")
    obj = ws.Range(rng) if rng else ws.UsedRange
    find = args["find"]
    replace = args.get("replace", None)
    look_in = args.get("look_in", -4163)
    match_case = args.get("match_case", False)
    look_at = 2 if args.get("match_partial", False) else 1
    if replace is not None:
        obj.Replace(What=find, Replacement=replace, LookAt=look_at, SearchOrder=1,
                    MatchCase=match_case)
        return {"replaced": find, "with": replace}
    found = obj.Find(What=find, LookIn=look_in, LookAt=look_at, MatchCase=match_case)
    if found and str(found) != "Nothing":
        return {"found": find, "at": f"{found.Row},{found.Column}"}
    return {"found": None}

@app.tool("wps_freeze",
    "冻结窗格",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "冻结到此单元格如 B2（不传=取消冻结）"}
        }
    })
def wps_freeze(args):
    cell = args.get("cell", "")
    et = _get_et()
    w = et.ActiveWindow
    if not cell:
        w.FreezePanes = False
        return {"frozen": False}
    w.ScrollRow = 1
    w.ScrollColumn = 1
    ws = _get_ws()
    ws.Range(cell).Select()
    w.FreezePanes = True
    return {"frozen": cell}

@app.tool("wps_protect",
    "工作表保护与取消保护 — unprotect=true 取消保护",
    input_schema={
        "type": "object",
        "properties": {
            "password": {"type": "string", "description": "设置密码（不传=无密码）"},
            "unprotect": {"type": "boolean", "description": "设为 true 取消保护，默认 false=启用保护"}
        }
    })
def wps_protect(args):
    ws = _get_ws()
    pw = args.get("password")
    if args.get("unprotect"):
        ws.Unprotect(pw or "")
        return {"protected": False}
    if pw:
        ws.Protect(pw)
    else:
        ws.Protect()
    return {"protected": True}

@app.tool("wps_group",
    "组合/分级显示（行或列分组）— ungroup=true 取消",
    input_schema={
        "type": "object",
        "properties": {
            "rows": {"type": "string", "description": "行范围如 2:10"},
            "cols": {"type": "string", "description": "列范围如 B:D"},
            "ungroup": {"type": "boolean", "description": "true=取消组合 / false=创建组合（默认 false）"}
        }
    })
def wps_group(args):
    ws = _get_ws()
    rows = args.get("rows", "")
    cols = args.get("cols", "")
    is_ungroup = args.get("ungroup", False)
    method = "Ungroup" if is_ungroup else "Group"
    if rows:
        getattr(ws.Range(f"A{rows}").EntireRow, method)()
        return {("grouped" if not is_ungroup else "ungrouped"): rows}
    if cols:
        getattr(ws.Range(f"{cols.split(':')[0]}1:{cols.split(':')[1]}1"), method)()
        return {("grouped" if not is_ungroup else "ungrouped"): cols}
    return {"error": "请指定 rows 或 cols"}

@app.tool("wps_page_setup",
    "页面设置（打印区域/边距/方向/纸张/页眉页脚等）",
    input_schema={
        "type": "object",
        "properties": {
            "orientation": {"type": "integer", "description": "1=纵向 2=横向"},
            "paper_size": {"type": "integer", "description": "纸张 1=信纸 9=A4 13=B5"},
            "fit_to_pages_wide": {"type": "integer", "description": "页宽"},
            "fit_to_pages_tall": {"type": "integer", "description": "页高"},
            "zoom": {"type": "integer", "description": "缩放比例 10-400"},
            "top_margin": {"type": "number", "description": "上边距（磅）"},
            "bottom_margin": {"type": "number", "description": "下边距"},
            "left_margin": {"type": "number", "description": "左边距"},
            "right_margin": {"type": "number", "description": "右边距"},
            "header_margin": {"type": "number", "description": "页眉边距"},
            "footer_margin": {"type": "number", "description": "页脚边距"},
            "center_horizontally": {"type": "boolean", "description": "水平居中"},
            "center_vertically": {"type": "boolean", "description": "垂直居中"},
            "print_area": {"type": "string", "description": "打印区域如 A1:Z100"},
            "print_titles_rows": {"type": "string", "description": "顶端标题行如 $1:$1"},
            "print_titles_cols": {"type": "string", "description": "左端标题列如 $A:$A"},
            "header": {"type": "string", "description": "页眉文字"},
            "footer": {"type": "string", "description": "页脚文字"},
        }
    })
def wps_page_setup(args):
    ws = _get_ws()
    ps = ws.PageSetup
    if "orientation" in args: ps.Orientation = args["orientation"]
    if "paper_size" in args: ps.PaperSize = args["paper_size"]
    if "fit_to_pages_wide" in args: ps.FitToPagesWide = args["fit_to_pages_wide"]
    if "fit_to_pages_tall" in args: ps.FitToPagesTall = args["fit_to_pages_tall"]
    if "zoom" in args: ps.Zoom = args["zoom"]
    if "top_margin" in args: ps.TopMargin = args["top_margin"]
    if "bottom_margin" in args: ps.BottomMargin = args["bottom_margin"]
    if "left_margin" in args: ps.LeftMargin = args["left_margin"]
    if "right_margin" in args: ps.RightMargin = args["right_margin"]
    if "header_margin" in args: ps.HeaderMargin = args["header_margin"]
    if "footer_margin" in args: ps.FooterMargin = args["footer_margin"]
    if "center_horizontally" in args: ps.CenterHorizontally = args["center_horizontally"]
    if "center_vertically" in args: ps.CenterVertically = args["center_vertically"]
    if "print_area" in args: ps.PrintArea = args["print_area"]
    if "print_titles_rows" in args: ws.PageSetup.PrintTitleRows = args["print_titles_rows"]
    if "print_titles_cols" in args: ws.PageSetup.PrintTitleColumns = args["print_titles_cols"]
    if "header" in args: ps.CenterHeader = args["header"]
    if "footer" in args: ps.CenterFooter = args["footer"]
    return {"page_setup": "ok"}

@app.tool("wps_chart",
    "创建图表",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "数据源区域如 A1:C10"},
            "type": {"type": "integer", "description": "图表类型 51=柱状图 4=折线图 1=面积图 57=条形图 5=饼图 -4169=散点图 65=折线图(带标记)  -4100=三维柱形 -4102=三维饼图 -4120=圆环图"},
            "title": {"type": "string", "description": "图表标题"},
            "x_title": {"type": "string", "description": "X轴标题"},
            "y_title": {"type": "string", "description": "Y轴标题"},
            "has_legend": {"type": "boolean", "description": "显示图例"}
        },
        "required": ["range"]
    })
def wps_chart(args):
    ws = _get_ws()
    rng = args["range"]
    chart_type = args.get("type", 51)
    cobj = ws.ChartObjects().Add(100, 50, 500, 300)
    cht = cobj.Chart
    cht.SetSourceData(ws.Range(rng))
    cht.ChartType = chart_type
    if "title" in args: cht.HasTitle = True; cht.ChartTitle.Text = args["title"]
    if "x_title" in args: cht.Axes(1).HasTitle = True; cht.Axes(1).AxisTitle.Text = args["x_title"]
    if "y_title" in args: cht.Axes(2).HasTitle = True; cht.Axes(2).AxisTitle.Text = args["y_title"]
    if "has_legend" in args: cht.HasLegend = args["has_legend"]
    return {"chart": rng, "type": chart_type}

def _walk_com(start, target):
    """Walk a dotted path like 'Range(\"A1\").Font.Bold' from start object"""
    parts = re.findall(r"(\w+)(?:\(((?:[^()]|\[[^\]]*\])*)\))?\.?", target)
    obj = start
    for name, raw_args in parts:
        if raw_args:
            args = []
            for m in re.finditer(r"""\[([^\]]*)\]|"([^"]*)"|'([^']*)'|(\d+\.?\d*)|(True|False)""", raw_args):
                arr, q1, q2, num, flag = m.groups()
                if arr is not None:
                    items = [int(x.strip()) if x.strip().isdigit() else x.strip() for x in arr.split(",") if x.strip()]
                    args.append(tuple(items))
                elif q1 is not None: args.append(q1)
                elif q2 is not None: args.append(q2)
                elif flag == "True": args.append(True)
                elif flag == "False": args.append(False)
                else: args.append(float(num) if "." in num else int(num))
            obj = getattr(obj, name)(*args)
        else:
            obj = getattr(obj, name)
    return obj

@app.tool("wps_com",
    "通用 COM 对象操作：读写任意属性、调用任意方法",
    input_schema={
        "type": "object",
        "properties": {
            "target": {"type": "string", "description": "对象路径，如 Range('A1').Font"},
            "prop": {"type": "string", "description": "属性名（与 value 配合读取或设置，如 'Bold'）"},
            "value": {"description": "属性值（传则写入属性，不传则读取）"},
            "call": {"type": "string", "description": "调用方法，如 'Copy' 'Clear' 'Delete'"}
        },
        "required": ["target"]
    })
def wps_com(args):
    target = args["target"]
    if target.startswith("Application."):
        obj = _walk_com(_get_et(), target)
    else:
        obj = _walk_com(_get_ws(), target)
    if "call" in args:
        method = getattr(obj, args["call"])
        result = method()
        return {"call": f"{target}.{args['call']}", "result": str(result) if result else None}
    if "prop" in args:
        prop = args["prop"]
        if "value" in args:
            setattr(obj, prop, args["value"])
            return {"set": f"{target}.{prop}", "value": args["value"]}
        val = getattr(obj, prop)
        try:
            val = str(val)
        except Exception:
            val = "<COM object>"
        return {"get": f"{target}.{prop}", "value": val}
    try:
        return {"result": str(obj)}
    except Exception:
        return {"result": "<COM object>"}

@app.tool("wps_batch",
    "批量执行多个 wps 操作（一次 MCP 往返，提高效率）",
    input_schema={
        "type": "object",
        "properties": {
            "steps": {
                "type": "array",
                "description": "操作列表，每步指定 tool 和 args",
                "items": {"type": "object"}
            }
        },
        "required": ["steps"]
    })
def wps_batch(args):
    ws = _get_ws()
    results = []
    for step in args["steps"]:
        t = step.get("tool", "")
        fn_name = t if t.startswith("wps_") else f"wps_{t}"
        fn = globals().get(fn_name)
        if fn and callable(fn):
            params = step.get("params") or step.get("args", {})
            results.append(fn(params))
            continue
        results.append({"step": step, "status": "unknown_tool"})
    return {"batch": results}

@app.tool("wps_close",
    "关闭当前工作簿（不保存），并清理临时文件【仅当用户明确要求时才调用，别自己关】")
def wps_close(args):
    global _wb, _orig_path
    if _wb is None:
        raise RuntimeError("没有打开的文件")
    tmp = _wb.FullName
    _wb.Close(SaveChanges=False)
    _wb = None
    _orig_path = None
    try:
        os.remove(tmp)
    except Exception:
        pass
    return {"closed": "已关闭" + (f"，临时文件已清理" if os.path.exists(tmp) else "")}

@app.tool("wps_save_as",
    "另存为",
    input_schema={
        "type": "object",
        "properties": {
            "path": {"type": "string", "description": "另存为路径，如 D:/test.xlsx"},
            "type": {"type": "integer", "description": "文件格式 51=xlsx 23=csv(UTF8) 56=csv(ANSI) 6=csv(Mac) -4158=PDF(当页) 0=xls"}
        },
        "required": ["path"]
    })
def wps_save_as(args):
    global _wb
    if _wb is None:
        raise RuntimeError("没有打开的文件")
    path = os.path.abspath(args["path"])
    ext = args.get("type", 51)
    _wb.SaveAs(path, ext)
    return {"saved_as": path}

@app.tool("wps_rename_sheet",
    "重命名工作表",
    input_schema={
        "type": "object",
        "properties": {
            "old_name": {"type": "string", "description": "当前工作表名"},
            "new_name": {"type": "string", "description": "新工作表名"}
        },
        "required": ["old_name", "new_name"]
    })
def wps_rename_sheet(args):
    ws = _get_ws()
    old = args["old_name"]
    new = args["new_name"]
    for j in range(1, _wb.Sheets.Count+1):
        s = _wb.Sheets.Item(j)
        if str(s.Name) == old:
            s.Name = new
            return {"renamed": old, "to": new}
    raise RuntimeError(f"找不到工作表: {old}")

@app.tool("wps_hide_sheet",
    "隐藏/取消隐藏工作表 — hide=false 显示",
    input_schema={
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "要操作的工作表名称"},
            "hide": {"type": "boolean", "description": "true=隐藏 / false=显示（默认 true）"}
        }
    })
def wps_hide_sheet(args):
    name = args["name"]
    hide = args.get("hide", True)
    for j in range(1, _wb.Sheets.Count+1):
        s = _wb.Sheets.Item(j)
        if str(s.Name) == name:
            s.Visible = 2 if hide else -1
            return {"sheet": name, "visible": not hide}
    raise RuntimeError(f"找不到工作表: {name}")

@app.tool("wps_read_formula",
    "读取公式文本（原样显示公式，非计算结果）",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "单元格如 C1"}
        },
        "required": ["cell"]
    })
def wps_read_formula(args):
    ws = _get_ws()
    cell = args["cell"]
    v = ws.Range(cell).Formula
    return {"cell": cell, "formula": str(v) if v else ""}

@app.tool("wps_delete_column",
    "删除指定列",
    input_schema={
        "type": "object",
        "properties": {
            "col": {"type": "string", "description": "列如 C"}
        },
        "required": ["col"]
    })
def wps_delete_column(args):
    ws = _get_ws()
    col = args["col"].upper()
    n = _col_to_num(col)
    ws.Columns(n).Delete()
    return {"deleted": col}

@app.tool("wps_hide_row",
    "隐藏/取消隐藏行 — hide=false 显示",
    input_schema={
        "type": "object",
        "properties": {
            "row": {"type": "integer", "description": "行号"},
            "hide": {"type": "boolean", "description": "true=隐藏 / false=显示（默认 true）"}
        }
    })
def wps_hide_row(args):
    ws = _get_ws()
    row = args["row"]
    hide = args.get("hide", True)
    ws.Rows(row).Hidden = hide
    return {"row": row, "hidden": hide}

@app.tool("wps_hide_column",
    "隐藏/取消隐藏列 — hide=false 显示",
    input_schema={
        "type": "object",
        "properties": {
            "col": {"type": "string", "description": "列如 C"},
            "hide": {"type": "boolean", "description": "true=隐藏 / false=显示（默认 true）"}
        },
        "required": ["col"]
    })
def wps_hide_column(args):
    ws = _get_ws()
    col = args["col"].upper()
    n = _col_to_num(col)
    hide = args.get("hide", True)
    ws.Columns(n).Hidden = hide
    return {"col": col, "hidden": hide}

@app.tool("wps_remove_duplicates",
    "删除重复值",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:C100"},
            "columns": {"type": "array", "description": "检查重复的列号数组（从1开始），如[1,2]"}
        },
        "required": ["range"]
    })
def wps_remove_duplicates(args):
    ws = _get_ws()
    rng = args["range"]
    cols = args.get("columns", None)
    obj = ws.Range(rng)
    if cols:
        obj.RemoveDuplicates(Columns=cols, Header=1)
    else:
        obj.RemoveDuplicates(Header=1)
    return {"removed_duplicates": rng}

@app.tool("wps_data_validation",
    "数据验证（设置/删除单元格输入限制—type=0 删除验证）",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:A10"},
            "type": {"type": "integer", "description": "验证类型 3=整数 4=小数 5=序列 6=日期 7=时间 8=文本长度 0=任意"},
            "formula1": {"type": "string", "description": "最小值/序列内容（如 '1,2,3' 或 '=A1:A10'）"},
            "formula2": {"type": "string", "description": "最大值"},
            "operator": {"type": "integer", "description": "运算符 1=介于 2=等于 3=大于 4=大于等于 5=小于 6=小于等于 7=不介于 8=不等于（默认1）"},
            "error_msg": {"type": "string", "description": "错误提示文字"}
        },
        "required": ["range"]
    })
def wps_data_validation(args):
    ws = _get_ws()
    rng = args["range"]
    r = args.get("type", 0)
    f1 = args.get("formula1", "")
    f2 = args.get("formula2", "")
    op = args.get("operator", 1)
    err = args.get("error_msg", "")
    dv = ws.Range(rng).Validation
    if r == 0:
        dv.Delete()
        return {"validation": "removed"}
    dv.Delete()
    type_map = {0:0, 3:1, 4:2, 5:3, 6:4, 7:5, 8:6, 1:1, 2:2}
    op_map = {1:1, 2:2, 3:3, 4:4, 5:5, 6:6, 7:7, 8:8}
    excel_type = type_map.get(r, r)
    excel_op = op_map.get(op, 1)
    dv.Add(Type=excel_type, AlertStyle=1, Operator=excel_op, Formula1=f1, Formula2=f2)
    if err:
        dv.ErrorMessage = err
        dv.ErrorTitle = "输入错误"
        dv.ShowError = True
    return {"validation": rng, "type": r}

@app.tool("wps_delete_chart",
    "删除工作表中的图表",
    input_schema={
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "图表名称（如 Chart 1，不传则删除全部）"}
        }
    })
def wps_delete_chart(args):
    ws = _get_ws()
    name = args.get("name", "")
    if name:
        ws.ChartObjects(name).Delete()
        return {"deleted": name}
    for c in ws.ChartObjects():
        c.Delete()
    return {"deleted": "all"}

@app.tool("wps_update_chart",
    "修改图表设置",
    input_schema={
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "图表名称如 Chart 1"},
            "chart_type": {"type": "integer", "description": "图表类型同 wps_chart"},
            "source_data": {"type": "string", "description": "数据源区域如 A1:C10"},
            "title": {"type": "string", "description": "图表标题"}
        },
        "required": ["name"]
    })
def wps_update_chart(args):
    ws = _get_ws()
    name = args["name"]
    cobj = ws.ChartObjects(name)
    cht = cobj.Chart
    if "chart_type" in args:
        cht.ChartType = args["chart_type"]
    if "source_data" in args:
        cht.SetSourceData(ws.Range(args["source_data"]))
    if "title" in args:
        cht.HasTitle = True
        cht.ChartTitle.Text = args["title"]
    return {"updated": name}

@app.tool("wps_add_comment",
    "添加批注",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "单元格如 B2"},
            "text": {"type": "string", "description": "批注内容"}
        },
        "required": ["cell", "text"]
    })
def wps_add_comment(args):
    ws = _get_ws()
    cell = args["cell"]
    text = args["text"]
    try:
        ws.Range(cell).Comment.Delete()
    except Exception:
        pass
    ws.Range(cell).AddComment(text)
    return {"comment": cell, "text": text}

@app.tool("wps_read_comment",
    "读取批注",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "单元格如 B2"}
        },
        "required": ["cell"]
    })
def wps_read_comment(args):
    ws = _get_ws()
    cell = args["cell"]
    try:
        text = ws.Range(cell).Comment.Text()
    except Exception:
        text = ""
    return {"cell": cell, "comment": str(text)}

@app.tool("wps_delete_comment",
    "删除批注",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "单元格如 B2（不传则删除全部批注）"}
        }
    })
def wps_delete_comment(args):
    ws = _get_ws()
    cell = args.get("cell", "")
    if cell:
        ws.Range(cell).Comment.Delete()
        return {"deleted": cell}
    ws.Cells.ClearComments()
    return {"deleted": "all"}

@app.tool("wps_add_hyperlink",
    "插入超链接",
    input_schema={
        "type": "object",
        "properties": {
            "cell": {"type": "string", "description": "单元格如 B2"},
            "address": {"type": "string", "description": "链接地址，如 https://www.baidu.com 或 Sheet2!A1"},
            "text": {"type": "string", "description": "显示文字（可选，默认用地址）"},
            "tip": {"type": "string", "description": "屏幕提示文字（可选）"}
        },
        "required": ["cell", "address"]
    })
def wps_add_hyperlink(args):
    ws = _get_ws()
    cell = args["cell"]
    addr = args["address"]
    text = args.get("text", addr)
    tip = args.get("tip", "")
    ws.Range(cell).Hyperlinks.Add(Anchor=ws.Range(cell), Address=addr, TextToDisplay=text, ScreenTip=tip)
    return {"hyperlink": cell, "address": addr}

@app.tool("wps_merge_columns",
    "文本合并：将多列内容按指定分隔符合并到一列",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "要合并的列区域如 A1:C10"},
            "delimiter": {"type": "string", "description": "分隔符（默认空字符串）"},
            "dest_cell": {"type": "string", "description": "目标单元格如 D1"}
        },
        "required": ["range", "dest_cell"]
    })
def wps_merge_columns(args):
    ws = _get_ws()
    rng = args["range"]
    delim = args.get("delimiter", "")
    dest = args["dest_cell"]
    cells = ws.Range(rng)
    r, c = cells.Rows.Count, cells.Columns.Count
    for i in range(1, r+1):
        parts = []
        for j in range(1, c+1):
            v = cells.Cells(i, j).Value
            if v is not None:
                parts.append(str(v))
        ws.Range(dest).Cells(i, 1).Value = delim.join(parts)
    return {"merged": rng, "to": dest}

@app.tool("wps_clear_conditional_format",
    "清除条件格式",
    input_schema={
        "type": "object",
        "properties": {
            "range": {"type": "string", "description": "区域如 A1:C10（不传则清除全表）"}
        }
    })
def wps_clear_conditional_format(args):
    ws = _get_ws()
    rng = args.get("range", "")
    if rng:
        ws.Range(rng).FormatConditions.Delete()
    else:
        ws.Cells.FormatConditions.Delete()
    return {"cleared": rng or "all"}

@app.tool("wps_delete_name",
    "删除命名区域",
    input_schema={
        "type": "object",
        "properties": {
            "name": {"type": "string", "description": "命名区域名称"}
        },
        "required": ["name"]
    })
def wps_delete_name(args):
    global _wb
    if _wb is None:
        raise RuntimeError("没有打开的文件")
    name = args["name"]
    try:
        _wb.Names(name).Delete()
    except Exception as e:
        raise RuntimeError(f"找不到命名区域: {name}")
    return {"deleted": name}

@app.tool("wps_run",
    "执行 Python 表达式（内置 ws/_wb/_et 对象），支持 print() 输出",
    input_schema={
        "type": "object",
        "properties": {
            "code": {"type": "string", "description": "Python 代码。内置变量：ws=当前表, _wb=当前工作簿, _et=WPS应用, _get_ws()=获取ws。用 eval 返回结果，多行语句用 print() 输出"}
        },
        "required": ["code"]
    })
def wps_run(args):
    code = args["code"]
    ws = _get_ws()
    g = dict(__builtins__=__builtins__ if isinstance(__builtins__, dict) else __builtins__.__dict__, ws=ws, _wb=_wb, _et=_get_et(), _get_ws=_get_ws)
    import io, sys as _sys
    buf = io.StringIO()
    old = _sys.stdout
    _sys.stdout = buf
    try:
        try:
            result = eval(code, g)
        except SyntaxError:
            exec(code, g)
            result = None
    finally:
        _sys.stdout = old
    out = buf.getvalue().strip()
    if out:
        return {"result": out}
    return {"result": str(result) if result is not None else None}

@app.tool("wps_help",
    "MCP 使用说明（本信息）",
    input_schema={
        "type": "object",
        "properties": {
            "topic": {"type": "string", "description": "可选：文件/读写/工作表/格式/数据/高级/枚举"}
        }
    })
def wps_help(args):
    topic = args.get("topic", "")
    help_texts = {
        "": """【WPS ET MCP 使用说明】
所有操作在临时副本上执行，wps_save 才写回原件。
直接 wps_close = 修改丢弃，原件不变。

常用流程：
  wps_open("path") → 各种操作 → wps_save → wps_close

📌 使用经验：
  - 排版数据用 wps_run 的 ws.Cells(r,c).Value = ...，比 wps_write 混写更稳
  - wps_find_replace 删制表符等特殊字符要加 match_partial=true
  - wps_count 的 * 通配符只匹配文本，数字单元格用 ">0"
  - 边框设 all=true 才出全框，只给 range 不生效
  - 字体颜色用 RGB 逗号分隔如 "255,0,0"
  - 多个操作合成一次 wps_batch 调用，速度快很多
  - 每次重启后如果 WPS 锁文件，先杀 et.exe 进程再 wps_open

查看具体分类：wps_help(topic="文件/读写/工作表/格式/数据/高级/枚举")
⚠ 重要：干完活别自己关文件！先问用户还有没有任务。
""",
        "文件": """文件操作：
  wps_open(path, sheet?)    打开文件（自动 temp 副本）
  wps_new(sheet_name?)      新建空白工作簿
  wps_save                  副本写回原件
  wps_close                 关闭副本不保存""",
        "读写": """读写数据：
  wps_read(range?)          读区域，默认 A1:T20（自动 strip 前后空白）
  wps_read_all(range?)      读全部（自动 strip 前后空白）
  wps_write(start_cell, data)  写入二维数组（大批量方便，但注意行号对齐）
  wps_update(cell, value)   改单个单元格
  wps_formula(cell, formula) 写公式 =SUM(A1:A10)
  wps_count(range, criteria) CountIf 统计（*通配符只匹配文本）
  wps_com(target, prop?, value?, call?)  万能 COM
  wps_run(code)             执行 Python（内置 ws/_wb/_et，ws.Cells 逐行写更精确）""",
        "工作表": """工作表管理：
  wps_list_sheets           列出所有表
  wps_switch_sheet(name)    切换
  wps_add_sheet(name?)      新建
  wps_copy_sheet(name?, new_name?)  复制
  wps_extract_sheet(range, from_sheet?, new_sheet?)  提取区域→新表
  wps_delete_sheet(name?)   删除""",
        "格式": """格式与样式：
  wps_format(range, merge?, unmerge?, row_height?, col_width?,
             font_bold?, font_size?, font_color?, fill_color?,
             highlight_duplicates?, data_bar?, color_scale?,
             auto_filter?)    合并/行高/列宽/字体/颜色/高亮重复/数据条/色阶/筛选
  wps_border(range, all?, left?, right?, top?, bottom?,
             inner?, outer?, color?, weight?, style?)  边框
  wps_freeze(cell?)         冻结窗格 (A2=首行)
  wps_protect(password?, unprotect?)  保护/取消""",
        "数据": """数据处理：
  wps_sort(range, key_column, desc?, key2?, order2?, key3?, order3?, has_header?)  排序
  wps_text_to_columns(range, delimiter?, tab?, semicolon?)  分列
  wps_find_replace(range?, find, replace?, look_in?, match_case?, match_partial?)  查找替换（match_partial=true 部分匹配）
  wps_count(range, criteria)  CountIf（*只匹配文本，数字用 ">0"）
  wps_insert(row?, col?, rows?) 插入行/列
  wps_delete_row(row)       删行
  wps_delete_column(col)    删列
  wps_remove_duplicates(range, columns?)  删重复值
  wps_data_validation(range, type?, formula1?, ...)  数据验证
  wps_group(rows?, cols?, ungroup?)  组合/取消""",
        "高级": """高级功能：
  wps_chart(range, type?, title?, x_title?, y_title?, has_legend?)  图表
  wps_page_setup(orientation?, paper_size?, margins?...)  页面设置
  wps_com(target, prop?, value?, call?)  万能 COM 调任意属性方法
  wps_run(code)             执行 Python（ws/_wb/_et 内置）""",
        "枚举": """常用 COM 枚举值：
  边框位置：左7 右10 上8 下9 内横12 内纵11
  线型：实线1 虚线-4115 双线-4119
  图表类型：柱状51 折线4 饼图5 面积1 散点-4169
  颜色计算：(B<<16)+(G<<8)+R
  冻结：_get_et().ActiveWindow.FreezePanes = True
  排序：升序1 降序2  表头有=1 无=2
  查找：公式-4123 值-4163 批注-4144""",
    }
    text = help_texts.get(topic, f"未知分类: {topic}\n可选: 文件/读写/工作表/格式/数据/高级/枚举")
    return {"help": text}

def _cleanup():
    global _et, _wb, _et_proc
    try:
        if _wb is not None:
            _wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        if _et is not None:
            _et.Quit()
    except Exception:
        pass
    try:
        if _et_proc is not None and _et_proc.poll() is None:
            _et_proc.terminate()
    except Exception:
        pass

if __name__ == "__main__":
    pythoncom.CoInitialize()
    try:
        app.run()
    finally:
        _cleanup()
        pythoncom.CoUninitialize()
