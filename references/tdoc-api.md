# 腾讯文档表格 API 调用要点

## 调用方式

```bash
python <tencentdocs.py> tdoc_call sheet-mcp <tool> '<json_args>'
```

`<tencentdocs.py>` 在插件缓存里，路径随版本变化：

```
~/.workbuddy/plugins/cache/workbuddy-builtin/tencent-docs-plugin/<版本>/skills/tencent-docs/tencentdocs.py
```

`read_tdoc.py` 会自动 glob 取最新版本，也可用 `MOJIAN_TDOC_SCRIPT` 指定。

**注意**：`--args` 这种写法不存在，JSON 是第 3 个位置参数。
用 `tdoc_schema sheet-mcp <tool>` 查参数定义，不要猜。

## 常用工具

### get_sheet_info

拿子表 ID 与行列数。

```json
{"file_id": "YOUR_FILE_ID"}
```

返回 `structuredContent.sheets[]`，每项有 `sheet_id / sheet_name / row_count / col_count / hidden`。
**不要传 sheet_id**，这个工具不支持该参数。

### get_cell_data

```json
{"file_id": "...", "sheet_id": "YOUR_SHEET_ID",
 "start_row": 3400, "end_row": 3880, "start_col": 0, "end_col": 26,
 "return_csv": false}
```

- 单次上限 **20000 单元格**：`(end_row-start_row+1) × (end_col-start_col+1) ≤ 20000`。
  27 列时每次最多 740 行，脚本里取 700 行一页。
- `return_csv=false` 返回结构化 `cells[]`，每项 `{row, col, string_value, value}`。
- **`col` 是 0 基**：col0=A 日期、col1=B 项目、col2=C 运营、col4=E 总数、col5=F 备注。
- 文本在 `string_value`，数值在 `value`。

### get_merged_cells

确认 C 列（运营）合并范围时用，`start_col=2, end_col=2`。
不要用"向下延续到下一个名字"去猜，末尾行可能没被合并覆盖。

## 日期块

A 列为 `2026/11/6 星期五` 这种格式，**只有块首行有值**（合并单元格）。
块尾 = 下一个块首行 − 1；今天的块通常就是表格最后一块，块尾取最后有数据的行。

`read_tdoc.py` 的 `norm_date()` 把各种写法归一成 `(年, 月, 日)` 元组再比对，
避开日期里月/日前导零不一致的问题（例如 `3/5` 与 `03-05`）。

## 输出与编码

- 输出重定向到文件时必须用 **Windows 路径**，`/tmp` 与 Windows Python 不互通。
- 中文输出建议 `sys.stdout.reconfigure(encoding='utf-8')`。
- 返回是 JSON-RPC 外壳：`{"jsonrpc":"2.0","result":{"content":[...], "structuredContent":{...}}}`
  （错误时是 `{"error": {"message": ...}}`）。
