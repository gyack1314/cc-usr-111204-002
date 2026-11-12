"""生成采集器测试用的无宏合成OOXML工作簿。

全部为虚构技术参数资料：无宏、无真实报价、无图片、无VBA。
运行：python3 -I tools/make_samples.py
输出目录：data/generated/（与既有手工样本分离）
"""
import os
import zipfile

CT='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/spec.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'''
ROOT_RELS='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
WB_RELS='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="spec" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/spec.xml"/><Relationship Id="str" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="style" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>{external}</Relationships>'''
EXT_REL='<Relationship Id="ext1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///never/opened/外部报价.xlsx" TargetMode="External"/>'

def workbook(date1904,sheets='<sheet name="燃机参数" sheetId="9" r:id="spec"/>'):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<workbookPr date1904="%d"/>%s<sheets>%s</sheets></workbook>'%(1 if date1904 else 0,'',sheets))

def write(path,parts,date1904,external=False,order=None):
    """parts: dict 成员名->文本；order给出ZIP物理成员顺序用于重排测试。"""
    full={'[Content_Types].xml':CT,'_rels/.rels':ROOT_RELS,
          'xl/workbook.xml':workbook(date1904),
          'xl/_rels/workbook.xml.rels':WB_RELS.format(external=EXT_REL if external else '')}
    full.update(parts)
    names=order if order else list(full)
    assert set(names)==set(full)
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for n in names:z.writestr(n,full[n])

SST='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="9"><si><t>性能总览</t></si><si><r><t>掺氢</t></r><r><t>比例</t></r></si><si><t>投运日期</t></si><si><t>负荷</t></si><si><t>效率</t></si><si><t>备注</t></si><si><t>布尔项</t></si><si><t>富文本</t></si><si><r><t>红</t></r><r><rPh t="拼音占位"/><t>外</t></r></si></sst>'''

# 样式：0常规 1百分比(10) 2内置日期(14) 3自定义日期 4自定义百分比 5常规
STYLES='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="0.00%"/></numFmts><cellXfs count="6"><xf numFmtId="0"/><xf numFmtId="10"/><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="0"/></cellXfs></styleSheet>'''

def sheet(rows_xml,extra=''):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      '<cols><col min="3" max="3" hidden="1" width="8"/></cols>'
      '<sheetData>%s</sheetData>%s</worksheet>'%(rows_xml,extra))

MERGES='<mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells>'
AUTOFILTER='<autoFilter ref="A2:F2"/>'

# 正文设计（行3-6）：
#  A: 百分比词法0.5（s=1，不乘100）；B: 日期序号（s=2/3，含1900虚构60）；
#  C: 布尔；D: 错误与空串；E: 稀疏尾列只在第6行出现；F: 公式（普通/无缓存/共享）
ROWS='''<row r="1"><c r="A1" t="s"><v>0</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="s"><v>2</v></c><c r="C2" t="s"><v>6</v></c><c r="D2" t="s"><v>5</v></c><c r="E2" t="s"><v>3</v></c><c r="F2" t="s"><v>4</v></c></row>
<row r="3"><c r="A3" s="1"><v>0.5</v></c><c r="B3" s="2"><v>44927</v></c><c r="C3" t="b"><v>1</v></c><c r="D3" t="e"><v>#DIV/0!</v></c><c r="F3"><f>A3*100</f><v>50</v></c></row>
<row r="4" hidden="1"><c r="A4" s="4"><v>0.125</v></c><c r="B4" s="2"><v>60</v></c><c r="C4" t="b"><v>0</v></c><c r="D4" t="inlineStr"><is><t></t></is></c><c r="F4"><f>A4*100</f></c></row>
<row r="5"><c r="A5" s="1"><v>0.25</v></c><c r="B5" s="3"><v>44927.5</v></c><c r="D5"><f t="shared" si="7" ref="D5:D6">A5+$B$5</f><v>1.25</v></c><c r="F5"><f>B5+1</f></c></row>
<row r="6"><c r="A6"><v>1E3</v></c><c r="B6" s="3"><v>0</v></c><c r="E6" t="s"><v>7</v></c><c r="D6"><f t="shared" si="7"/><v>2</v></c><c r="F6" t="s"><v>8</v></c></row>'''
CONFIG={
 'workbook':'parameters-rich.xlsx','sheet':'燃机参数',
 'header_first':1,'header_last':2,'body_first':3,'body_last':6,
 'date_style_ids':[2,3],'row_cache_limit':20}

def main():
    out=os.path.join(os.path.dirname(__file__),'..','data','generated')
    os.makedirs(out,exist_ok=True)
    base={'xl/sharedStrings.xml':SST,'xl/styles.xml':STYLES,
          'xl/worksheets/spec.xml':sheet(ROWS,MERGES+AUTOFILTER)}
    # 1) 富样本（1900纪元，含虚构序号60、隐藏列、筛选、稀疏尾列、共享公式）
    write(os.path.join(out,CONFIG['workbook']),base,date1904=False,external=True)
    # 2) 同内容1904纪元
    write(os.path.join(out,'parameters-rich-1904.xlsx'),base,date1904=True)
    # 3) 乱序共享公式：从属格D6物理上出现在基格D5之前（行节点乱序）
    rows6_before5=ROWS.replace(
      '<row r="5">','@@R5@@').replace('<row r="6">','@@R6@@')
    import re
    m5=re.search(r'@@R5@@.*?</row>',rows6_before5).group(0)
    m6=re.search(r'@@R6@@.*?</row>',rows6_before5).group(0)
    swapped=rows6_before5.replace(m5+m6,m6+m5).replace('@@R5@@','<row r="5">').replace('@@R6@@','<row r="6">')
    write(os.path.join(out,'parameters-shared-unordered.xlsx'),
          {**base,'xl/worksheets/spec.xml':sheet(swapped,MERGES)},date1904=False)
    # 4) 共享公式缺基：只有从属格si=7，无ref基格
    no_master=ROWS.replace('<c r="D5"><f t="shared" si="7" ref="D5:D6">A5+$B$5</f><v>1.25</v></c>',
                           '<c r="D5"><f t="shared" si="7"/><v>1.25</v></c>')
    write(os.path.join(out,'parameters-shared-no-master.xlsx'),
          {**base,'xl/worksheets/spec.xml':sheet(no_master,MERGES)},date1904=False)
    # 5) 共享公式多基冲突：si=7出现两个带ref且公式不同的基格
    conflict=ROWS.replace('<c r="D6"><f t="shared" si="7"/><v>2</v></c>',
        '<c r="D6"><f t="shared" si="7" ref="D6:D6">A6*9</f><v>2</v></c>')
    write(os.path.join(out,'parameters-shared-conflict.xlsx'),
          {**base,'xl/worksheets/spec.xml':sheet(conflict,MERGES)},date1904=False)
    # 6) 不支持语法：结构化引用/动态数组，保留原式与原因
    unsupported=ROWS.replace('<c r="F3"><f>A3*100</f><v>50</v></c>',
        '<c r="F3"><f>Table1[负荷]*2</f><v>50</v></c>')
    write(os.path.join(out,'parameters-unsupported-formula.xlsx'),
          {**base,'xl/worksheets/spec.xml':sheet(unsupported,MERGES)},date1904=False)
    import json
    for name in ('parameters-rich.xlsx','parameters-rich-1904.xlsx',
                 'parameters-shared-unordered.xlsx','parameters-shared-no-master.xlsx',
                 'parameters-shared-conflict.xlsx','parameters-unsupported-formula.xlsx'):
        cfg=dict(CONFIG);cfg['workbook']=name
        with open(os.path.join(out,name.replace('.xlsx','.json')),'w',encoding='utf8') as f:
            json.dump(cfg,f,ensure_ascii=False,indent=2)
    print('generated in',os.path.normpath(out))

if __name__=='__main__':main()
