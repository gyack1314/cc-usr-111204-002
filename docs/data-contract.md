# 工作簿资料

支持无宏OOXML工作簿（.xlsx），包成员路径按关系URI相对所在成员解析，TargetMode=External只保留外部关系，不连接。工作表与共享字符串均通过`workbook.xml`及其`_rels`关系定位，不假设sheet1.xml就是第一页。共享公式的si仅在单个工作表中有效，基公式f带ref区域，其他从属f只带si，乱序XML不改变对应关系。

## 采集范围与表头

采集范围完全由配置给定，不猜测表格边界：`sheet`（名称）、`header_first/header_last`（表头行区间）、`body_first/body_last`（正文行区间）、`date_style_ids`（允许转日期的cellXfs样式索引，即单元格`s`属性）、`row_cache_limit`（有界行缓存容量）。

合并区域只以左上角作为表头来源，形成逐列有序表头路径（`header_path`），合并覆盖列继承左上角条目但不把合并值写成正文每格真实值；重复路径以来源单元格列坐标保留区分。隐藏行列与autoFilter状态只记录、不默认排除。

## 单元格状态（互不混淆）

- `missing-cell`：配置正文区间内该坐标没有c元素。
- `empty-inline-string`：inlineStr的`<is><t></t></is>`，值为空字符串。
- `formula-no-cache`：有f元素但无v缓存（`cache_state=no-cache`）。
- `error`：`t="e"`显式错误，保留错误词法（如`#N/A`、`#DIV/0!`）。
- 另有`shared-string`/`inline-string`/`boolean`/`number`/`formula-string`等；共享字符串索引越界为`shared-index-invalid`。

数值同时保留原词法串`lexical`与`Decimal`文本`decimal`（如`1E3`与`1E+3`）。百分比只由样式分类`classification=percent`和`percent=true`说明，底层0.5不改成50。

## 日期与纪元

日期仅对`date_style_ids`列出的样式转换，按workbook的1900或1904纪元（`workbookPr@date1904`）。1900纪元下序号60（含小数）标`date_note=fictional-1900-02-29`而非合法公历日；1904纪元无此虚构日，序号60正常换算。未配置为日期的数值即使内置numFmt像日期也只标`date-unconfigured`，不擅自转换。风格表仍保存numFmtId等原始样式引用。

## 公式

公式不求值、不抓外部工作簿。支持词法：A1单格/矩形范围、可带单引号转义表名、绝对/相对/混合坐标、数字、双引号字符串、算术比较运算、括号、逗号和普通函数名。共享公式展开相对、绝对和混合A1引用；双引号字符串与引号工作表名内的字符不当引用；带表名引用不平移。结构化表引用、外部工作簿引用、动态数组专用语法（`@`）、R1C1样式拒绝展开但保留原式与`expand_reason`。相对引用超出Excel行列边界记录展开失败，不截断。

共享公式组收集基单元格及si，基公式可能物理上出现在从属格之后（独立索引阶段解决前向引用）。组缺基诊断`shared-formula-missing-master`，多基冲突诊断`shared-formula-conflicting-masters`，基格不在ref左上角诊断`shared-formula-master-not-at-ref-origin`。公式缓存与公式类型独立保留（`cache_state`为`cached`/`no-cache`/`not-a-formula`）。

## 输出与内存

逐格记录包含：包成员`package_member`、sheet关系ID `sheet_rid`、坐标`coordinate`、表头来源`header_path`、原类型`origin_type`、样式`style`（xf_id/raw_s/num_fmt_id/classification）、原公式`formula.raw`、展开公式`formula.expanded`及不可展开原因、缓存状态`cache_state`、隐藏标记与值/状态。

工作表XML用iterparse处理两遍：第一遍只建轻量索引，第二遍以容量`row_cache_limit`的有界堆缓存按坐标顺序产出，不把整张工作表树常驻内存。物理行节点乱序距离需落在缓存窗口内；坐标身份始终取自`r`属性，不随节点顺序改变。

`python -m source_data.cli collect <config.json> <out.jsonl>`流式输出，首行为meta，其后每行一条逐格记录。重排ZIP成员顺序不改变采集语义；采集过程只读不改，原文件字节保持不变。

## 样本

`data/parameters*.xlsx`为手工构造的最小样本（正常、乱序共享公式、1904纪元）。`data/generated/`由`tools/make_samples.py`重新生成，覆盖两套纪元、乱序共享公式、缺基/冲突基、不支持公式语法、合并表头、稀疏尾列、隐藏列、筛选、富文本rPh、布尔、错误与外部关系。所有资料均为虚构技术参数，不含宏、真实报价、图片或选中代码。

当前`domain.summarize`只列ZIP包成员，不读取工作表单元格；逐格采集见`source_data.xlsx`。
