# wps-office-api

WPS Office 自动化操作 MCP Server（ET 表格）。提供 46 个工具，覆盖 78 项测试点中的 74 项。

## 环境

```bash
pip install pywin32
```

**注意**: 需要安装 WPS Office（64位电信版），Python 与 WPS 架构必须一致（同为 64 位）。

## MCP Server

`et/wps_et_server.py` 通过 [MCP 协议](https://modelcontextprotocol.io) 提供 stdio 接口，AI Agent 可直接读写改删 WPS 表格。

### 使用方式

在 `mimocode.jsonc` 中添加：

```jsonc
{
  "$schema": "https://mimo.xiaomi.com//config.json",
  "mcp": {
    "wps-et": {
      "type": "local",
      "command": ["E:/python/python.exe", "-X", "utf8", "E:/mimo/wps-office-api-mcp/et/wps_et_server.py"],
      "enabled": true,
      "timeout": 60000
    }
  }
}
```

### 启动机制

Server 自动启动 `et.exe` 进程并通过 `GetActiveObject` 连接 COM，无需手动打开 WPS。

## 测试覆盖率

| 分类 | 通过 | 说明 |
|------|------|------|
| 一、工作簿管理（6项） | 6/6 | 创建/打开/信息/保存/另存为/关闭 |
| 二、工作表操作（8项） | 8/8 | 列出/创建/删除/重命名/切换/复制/隐藏/获取内容 |
| 三、单元格读写（10项） | 10/10 | 单格/范围/文本/数字/日期/批量/公式值/公式文本/命名区域/自动trim |
| 四、公式与计算（6项） | 6/6 | SUM/IF/VLOOKUP/自动重算/复制引用/错误处理 |
| 五、格式设置（8项） | 8/8 | 字体/背景/数字格式/边框/对齐/合并/行高列宽 |
| 六、行列操作（6项） | 6/6 | 插入删除行/列/隐藏行/列 |
| 七、数据管理（6项） | 6/6 | 排序/筛选/清除筛选/删除重复/数据验证 |
| 八、图表（4项） | 4/4 | 创建/修改类型/更新数据源/删除 |
| 九、透视表（3项） | ❌ 3/3 | WPS COM 不支持 PivotCaches.Create（与 Office 不同） |
| 批注/超链接（4项） | 4/4 | 添加/读取/删除批注/超链接 |
| 条件格式（2项） | 1/2 | 清除条件格式 ✅ / 设置条件格式 ❌ |
| 查找替换（2项） | 2/2 | 查找/替换 |
| 数据分列合并（2项） | 2/2 | 分列 / 合并列 |
| 保护（2项） | 2/2 | 保护/取消保护 |
| 分组（2项） | 2/2 | 创建分组/取消分组 |
| 页面设置（2项） | 2/2 | 打印区域/页边距方向 |
| 命名区域（1项） | 1/1 | 删除命名区域 |
| 冻结窗格（1项） | 1/1 | 冻结/取消冻结 |

**总计**: 78 项中通过 74 项（94.9%）。透视表 3 项因 WPS COM 限制无法支持。

## 未完成

| # | 功能 | 说明 |
|---|------|------|
| 1 | 设置条件格式 | `FormatConditions.Add` 参数适配待排查 |
| 2 | 多列排序 | 当前仅支持单列，需扩展 `wps_sort` |

## 工具列表（46个）

| 工具 | 功能 |
|------|------|
| `wps_open` | 打开 xlsx（副本操作） |
| `wps_new` | 新建空白工作簿 |
| `wps_save` | 保存修改回原件 |
| `wps_save_as` | 另存为（xlsx/csv/pdf 等格式） |
| `wps_close` | 关闭工作簿 |
| `wps_list_sheets` | 列出所有工作表 |
| `wps_add_sheet` | 新建工作表 |
| `wps_delete_sheet` | 删除工作表 |
| `wps_rename_sheet` | 重命名工作表 |
| `wps_switch_sheet` | 切换工作表 |
| `wps_copy_sheet` | 复制工作表 |
| `wps_hide_sheet` | 隐藏/显示工作表 |
| `wps_extract_sheet` | 提取区域内容到新表 |
| `wps_read` | 读取区域数据 |
| `wps_read_all` | 读取全部数据 |
| `wps_read_formula` | 读取公式文本 |
| `wps_write` | 写入二维数组 |
| `wps_update` | 修改单个单元格 |
| `wps_delete_row` | 删除指定行 |
| `wps_delete_column` | 删除指定列 |
| `wps_insert` | 插入行/列 |
| `wps_hide_row` | 隐藏/显示行 |
| `wps_hide_column` | 隐藏/显示列 |
| `wps_sort` | 排序（单列） |
| `wps_find_replace` | 查找替换（支持 `match_partial` 部分匹配） |
| `wps_remove_duplicates` | 删除重复值 |
| `wps_data_validation` | 数据验证 |
| `wps_text_to_columns` | 分列 |
| `wps_merge_columns` | 文本合并 |
| `wps_formula` | 写入公式 |
| `wps_format` | 格式（字体/颜色/合并/筛选/行高列宽） |
| `wps_border` | 边框设置 |
| `wps_freeze` | 冻结窗格 |
| `wps_protect` | 工作表保护 |
| `wps_group` | 组合/取消组合 |
| `wps_page_setup` | 页面设置 |
| `wps_chart` | 创建图表 |
| `wps_update_chart` | 修改图表 |
| `wps_delete_chart` | 删除图表 |
| `wps_add_comment` | 添加批注 |
| `wps_read_comment` | 读取批注 |
| `wps_delete_comment` | 删除批注 |
| `wps_add_hyperlink` | 插入超链接 |
| `wps_clear_conditional_format` | 清除条件格式 |
| `wps_delete_name` | 删除命名区域 |
| `wps_com` | 通用 COM 操作 |

## 文件

| 文件 | 说明 |
|------|------|
| `et/wps_et_server.py` | MCP Server 主程序（46个工具） |
| `et/mcp_base.py` | 轻量 MCP stdio 协议框架 |
| `et/et_crud.py` | 增删改查演示 |
| `et/et_batch.py` | 批量导入 + 公式计算 |
| `et/et_format.py` | 格式美化 + 图表示例 |
