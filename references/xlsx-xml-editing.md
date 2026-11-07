# 直接改 xlsx 内部 XML 的踩坑清单

本技能不依赖 openpyxl，全部直接操作 `xl/worksheets/sheetN.xml`，其余 zip 条目字节级保留
（openpyxl 重新保存会丢打印设置、条件格式、列宽等）。

## 铁律

**任何改 XML 的脚本，打包前必须 `ET.fromstring()` 校验，校验不过就不许写盘。**

曾因为属性拼错产出非法 XML，把 D 盘原表写坏过一次（靠备份恢复）。改表前先备份。

## 单元格正则必须原子二选一

```python
CELL = re.compile(r'<c\b[^>]*\br="B%d"[^>]*/>|<c\b[^>]*\br="B%d"[^>]*>.*?</c>' % (n, n), re.S)
```

**自闭合分支必须写在前面。**

- 写成 `<c r="B20"[^>]*(?:/>|>.*?</c>)` 会把自闭合单元格结尾的 `/` 吃掉，转而吞掉后面整个单元格
  （曾因此丢失 D20 的「默认部门」）。
- 写成 `<c\b[^>]*/>|<c\b[^>]*>.*?</c>` 却不用 `r="..."` 锚定，遇到 `<c r="B3" s="8"/>` 会与后一个单元格粘连。

## 重建单元格绝不能复用原 attrs

把原 `<c r="B2" s="8"/>` 的属性串整段带进新标签，会产出 `<c r="B2" r="B2" s="8"//>`：
r 重复 + `/` 错位，非法 XML。

正确做法：**只提取 `s="N"` 样式，ref 重新生成**。

```python
def style_of(cell):
    m = re.search(r'\bs="(\d+)"', cell)
    return ' s="%s"' % m.group(1) if m else ''

new = '<c r="B%d"%s><v>%d</v></c>' % (n, style_of(old), qty)
```

## 共享字符串表要用 ElementTree

不能拿正则 `<si>.*?</si>` 数，空 `<si/>` 会导致下标错位。

```python
SST = [''.join(t.text or '' for t in si.iter(NS + 't')) for si in ET.parse(...).getroot()]
```

## inlineStr 也是合法单元格

`add_projects.py` 插入的新行是 `<c ... t="inlineStr"><is><t>值</t></is></c>`，
按值匹配项目名时要能解析这种格式，不能只认 `<v>` 共享字符串索引。

## 插行后要修的东西

1. 下面所有行的 `row r` 与单元格 `ref` 都要 +1（`add_projects.shift_rows`）；
2. `<dimension ref="A1:E145">` 要同步；
3. 必须用 `check_insert.py` 自检「除新增行外原始数据逐行原序保留」。

## 写数值前先验行号

用 A 列项目名校验目标行，防止块结构变化导致错位写入：

```python
actual = read_a_text(xml, sst, n)
if actual.strip() != expect: raise SystemExit('A 列校验失败')
```

`apply_plan.py` 已内置这一步。

## 工作表解析不要写死 sheetN.xml

用 `xl/workbook.xml` + `xl/_rels/workbook.xml.rels` 按表名解析，拿到真实 target。
