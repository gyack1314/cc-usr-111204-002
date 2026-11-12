"""采集器测试：关系定位、稀疏坐标、类型状态、纪元、共享公式与ZIP重排不变性。

合成资料由tools/make_samples.py生成；既有data/*.xlsx为手工样本。
所有Python读取均以-I隔离运行（见README），测试内只用标准库。
"""
import hashlib
import io
import json
import unittest
import zipfile
from pathlib import Path
from source_data.io import load,asset
from source_data.xlsx import extract,open_workbook,ExtractError
from source_data.coords import col_to_index,col_to_letters,split_ref,split_range
from source_data.formulas import translate_shared,unsupported_reasons

DATA=Path('data')
GEN=DATA/'generated'

def cfg(name):return load(GEN/name)[0]
def raw_of(config,root=GEN):return asset(root,config['workbook'])[0]
def index(records,coord):return next(r for r in records if r['coordinate']==coord)

class CoordsTests(unittest.TestCase):
    def test_roundtrip(self):
        for letters,n in (('A',1),('Z',26),('AA',27),('XFD',16384)):
            self.assertEqual(col_to_index(letters),n)
            self.assertEqual(col_to_letters(n),letters)
        self.assertEqual(split_ref('XFD1048576'),(1048576,16384))
        self.assertEqual(split_range('A1:C3'),(1,1,3,3))
        with self.assertRaises(ValueError):split_ref('a1')
        with self.assertRaises(ValueError):col_to_letters(0)

class FormulaTests(unittest.TestCase):
    def test_relative_absolute_mixed(self):
        f='A3+$B$3+SUM(A$1:$C3)'
        self.assertEqual(translate_shared(f,'D3','D4')[0],
                         'A4+$B$3+SUM(A$1:$C4)')
        self.assertEqual(translate_shared(f,'D3','F8')[0],
                         'C8+$B$3+SUM(C$1:$C8)')
    def test_string_and_quoted_sheet_not_reference(self):
        f='\'掺氢 比例\'!A1+"B2"+"含""引号"'
        out,reason=translate_shared(f,'D3','F8')
        self.assertIsNone(reason);self.assertEqual(out,f)
    def test_unsupported_kept_verbatim(self):
        for bad in ('Table1[负荷]*2','A1+@B2','R[-1]C[2]'," '[外部]b.xlsx'!Sheet1!A1"):
            out,reason=translate_shared(bad,'D3','D4')
            self.assertEqual(out,bad);self.assertIsNotNone(reason)
    def test_out_of_bounds(self):
        out,reason=translate_shared('B1','A1','XFD1')
        self.assertIsNotNone(reason);self.assertIn('越界',reason);self.assertEqual(out,'B1')
    def test_r1c1_guard(self):
        self.assertEqual(unsupported_reasons('$C3+SUM(A1)'),[])
        self.assertTrue(unsupported_reasons('R5C2'))

class ExtractionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.configs={p.stem:load(p)[0] for p in GEN.glob('*.json')}
        cls.records={};cls.metas={}
        for stem,c in cls.configs.items():
            meta,recs=extract(raw_of(c),c)
            cls.records[stem]=recs;cls.metas[stem]=meta

    def test_relationships_not_sheet1_assumption(self):
        meta=self.metas['parameters-rich']
        self.assertEqual(meta['sheet_part'],'xl/worksheets/spec.xml')
        self.assertEqual(meta['sheet_rid'],'spec')
        self.assertEqual(meta['shared_strings_part'],'xl/sharedStrings.xml')
        self.assertTrue(meta['external_rels'])
        self.assertEqual(meta['external_rels'][0]['rid'],'ext1')

    def test_distinct_states(self):
        recs=self.records['parameters-rich']
        self.assertEqual(index(recs,'E3')['state'],'missing-cell')
        self.assertEqual(index(recs,'D4')['state'],'empty-inline-string')
        self.assertEqual(index(recs,'F4')['state'],'formula-no-cache')
        self.assertEqual(index(recs,'F4')['cache_state'],'no-cache')
        self.assertEqual(index(recs,'D3')['state'],'error')
        self.assertEqual(index(recs,'D3')['value'],'#DIV/0!')
        self.assertEqual(index(recs,'C3')['state'],'boolean')
        self.assertIs(index(recs,'C3')['value'],True)

    def test_percent_lexical_and_decimal_unchanged(self):
        a3=index(self.records['parameters-rich'],'A3')
        self.assertEqual(a3['value']['lexical'],'0.5')
        self.assertEqual(a3['value']['decimal'],'0.5')
        self.assertTrue(a3['value']['percent'])
        self.assertEqual(a3['style']['classification'],'percent')
        a6=index(self.records['parameters-rich'],'A6')
        self.assertEqual(a6['value']['lexical'],'1E3')
        self.assertEqual(a6['value']['decimal'],'1E+3')

    def test_merged_header_path_and_no_body_fill(self):
        recs=self.records['parameters-rich']
        f3=index(recs,'F3')
        self.assertEqual([p['value'] for p in f3['header_path']],['性能总览','效率'])
        self.assertEqual([p['source'] for p in f3['header_path']],['A1','F2'])
        # 合并区只作表头来源，正文E3仍是缺失格，不被填成合并值
        self.assertEqual(index(recs,'E3')['value'],None)
        self.assertEqual(self.metas['parameters-rich']['merged_areas'],['A1:F1'])

    def test_hidden_and_filter_kept(self):
        meta=self.metas['parameters-rich'];recs=recs=self.records['parameters-rich']
        self.assertEqual(meta['hidden_columns'],[3])
        self.assertEqual(meta['autofilter'],'A2:F2')
        self.assertTrue(index(recs,'A4')['row_hidden'])
        self.assertTrue(index(recs,'C3')['col_hidden'])

    def test_sparse_tail_column(self):
        recs=self.records['parameters-rich']
        self.assertEqual(index(recs,'E3')['state'],'missing-cell')
        self.assertEqual(index(recs,'E6')['state'],'shared-string')
        self.assertEqual(index(recs,'E6')['value'],'富文本')
        # rPh拼音占位不进正文
        self.assertEqual(index(recs,'F6')['value'],'红外')
        self.assertEqual(self.metas['parameters-rich']['total_columns'],6)

    def test_epoch_and_fictional_date(self):
        r1900=index(self.records['parameters-rich'],'B4')
        self.assertIsNone(r1900['value']['date_iso'])
        self.assertEqual(r1900['value']['date_note'],'fictional-1900-02-29')
        r1904=index(self.records['parameters-rich-1904'],'B4')
        self.assertEqual(r1904['value']['date_iso'],'1904-03-01')
        self.assertIsNone(r1904['value']['date_note'])
        self.assertEqual(index(self.records['parameters-rich'],'B3')['value']['date_iso'],
                         '2023-01-01')
        self.assertEqual(index(self.records['parameters-rich-1904'],'B3')['value']['date_iso'],
                         '2027-01-02')

    def test_unordered_shared_formula_matches_ordered(self):
        unordered=self.records['parameters-shared-unordered']
        ordered=self.records['parameters-rich']
        self.assertEqual([r['coordinate'] for r in unordered],
                         [r['coordinate'] for r in ordered])
        d6=index(unordered,'D6')
        self.assertEqual(d6['formula']['expanded'],'A6+$B$5')
        self.assertEqual(d6['formula']['master_cell'],'D5')
        self.assertEqual(self.metas['parameters-shared-unordered']['diagnostics'],[])

    def test_missing_and_conflicting_master(self):
        nm=self.metas['parameters-shared-no-master']
        self.assertEqual(nm['diagnostics'][0]['code'],'shared-formula-missing-master')
        cf=self.metas['parameters-shared-conflict']
        self.assertEqual(cf['diagnostics'][0]['code'],
                         'shared-formula-conflicting-masters')

    def test_origin_fields_present(self):
        r=index(self.records['parameters-rich'],'A3')
        for key in ('package_member','sheet_rid','coordinate','header_path','origin_type',
                    'style','formula','cache_state','state','value','row_hidden','col_hidden'):
            self.assertIn(key,r)

class ZipInvarianceTests(unittest.TestCase):
    def test_member_reorder_keeps_semantics_and_bytes(self):
        c=load(GEN/'parameters-rich.json')[0]
        path=GEN/c['workbook']
        before=path.read_bytes()
        with zipfile.ZipFile(io.BytesIO(before)) as z:
            infos=z.infolist()
            chunks={i.filename:z.read(i.filename) for i in infos}
        buf=io.BytesIO()
        order=[i.filename for i in infos][::-1]
        with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
            for n in order:z.writestr(n,chunks[n])
        reordered=buf.getvalue()
        meta1,recs1=extract(before,c)
        meta2,recs2=extract(reordered,c)
        self.assertEqual(
            json.dumps(recs1,ensure_ascii=False,sort_keys=True),
            json.dumps(recs2,ensure_ascii=False,sort_keys=True))
        self.assertEqual(meta1['sheet_part'],meta2['sheet_part'])
        # 原始文件字节未被采集过程改动
        self.assertEqual(path.read_bytes(),before)
        self.assertEqual(hashlib.sha256(before).hexdigest(),
                         hashlib.sha256(path.read_bytes()).hexdigest())

class BoundedCacheTests(unittest.TestCase):
    def test_cache_limit_one_still_orders_unordered(self):
        c=load(GEN/'parameters-shared-unordered.json')[0]
        c=dict(c);c['row_cache_limit']=1
        with open_workbook(raw_of(c),c) as wb:
            coords=[r['coordinate'] for r in wb.records()]
        # 物理乱序距离超过窗口1时按到达顺序输出，但坐标身份仍正确且无丢失
        self.assertEqual(sorted(coords),sorted(set(coords)))
        self.assertEqual(len(coords),24)

    def test_invalid_config_rejected(self):
        c=load(GEN/'parameters-rich.json')[0]
        for bad in ('header_first','body_first'):
            d=dict(c);d[bad]=0
            with self.assertRaises((ExtractError,ValueError)):extract(raw_of(d),d)
        d=dict(c);d['header_last']=0
        with self.assertRaises((ExtractError,ValueError)):extract(raw_of(d),d)
        d=dict(c);d['row_cache_limit']=0
        with self.assertRaises((ExtractError,ValueError)):extract(raw_of(d),d)
        d=dict(c);d['sheet']='不存在'
        with self.assertRaises(ExtractError):extract(raw_of(d),d)

class LegacySamplesTests(unittest.TestCase):
    def test_handcrafted_samples(self):
        for jf in ('normal.json','forward-formula.json','date-1904.json'):
            c=load(DATA/jf)[0]
            meta,recs=extract(asset(DATA,c['workbook'])[0],c)
            self.assertEqual(meta['sheet_rid'],'spec')
            self.assertEqual(len(recs),(c['body_last']-c['body_first']+1)*meta['total_columns'])
            d5=index(recs,'D5')
            self.assertEqual(d5['state'],'formula-no-cache')
            self.assertEqual(d5['cache_state'],'no-cache')

if __name__=='__main__':unittest.main()
