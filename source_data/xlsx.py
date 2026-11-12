"""按OOXML包关系采集指定工作表的逐格来源记录。

边界约定：
- 采集范围完全由配置的sheet名称、表头行区间和正文行区间决定，不猜测表格边界。
- 工作表与共享字符串通过workbook及其rels关系定位，不假设sheet1.xml就是第一页。
- 工作表XML用iterparse处理两遍：第一遍只建轻量索引（合并区、隐藏列、筛选、
  共享公式基格、表头快照明细），第二遍以有界行缓存（堆，容量row_cache_limit）
  按坐标顺序产出逐格记录，不把整张工作表树常驻内存。
- 物理行节点乱序距离需落在缓存窗口内；窗口外乱序无法在有界内存下重排，
  坐标身份始终取自单元格r属性，不随节点顺序改变。
"""
import heapq
import io
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from decimal import Decimal,InvalidOperation
from .coords import split_ref,join_ref,split_range
from .formulas import translate_shared,unsupported_reasons
from .io import integer

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKGNS='http://schemas.openxmlformats.org/package/2006/relationships'
def q(tag):return '{%s}%s'%(NS,tag)

BUILTIN_DATE_FMTS={14,15,16,17,22,27,30,36,50,57}
BUILTIN_TIME_FMTS={18,19,20,21,45,46,47}
EPOCH_1900_ORDINAL=date(1899,12,30).toordinal()
EPOCH_1904_ORDINAL=date(1904,1,1).toordinal()

class ExtractError(ValueError):pass

def _members(z):
    seen=set()
    for info in z.infolist():
        if info.filename in seen:raise ExtractError('工作簿包成员重名:'+info.filename)
        seen.add(info.filename)
    return seen

def _rels(xml_bytes):
    root=ET.fromstring(xml_bytes)
    rels={}
    for r in root.findall('{%s}Relationship'%PKGNS):
        rid=r.get('Id')
        if rid in rels:raise ExtractError('关系ID重复:'+rid)
        rels[rid]={'target':r.get('Target'),'mode':r.get('TargetMode'),'type':r.get('Type')}
    return rels

def _resolve(base_part,target):
    """按OPC相对URI规则解析包内目标路径，返回规范化成员名。"""
    parts=(target.lstrip('/').split('/') if target.startswith('/')
           else base_part.split('/')[:-1]+target.split('/'))
    stack=[]
    for p in parts:
        if p in ('','.'):continue
        if p=='..':
            if not stack:raise ExtractError('关系路径越界:'+target)
            stack.pop()
        else:stack.append(p)
    return '/'.join(stack)

def _rels_part_of(part):
    head,_,tail=part.rpartition('/')
    return (head+'/' if head else '')+'_rels/'+tail+'.rels'

def _si_text(si_el):
    """共享字符串富文本拼接；rPh拼音提示节点不是正文，必须排除。"""
    out=[]
    for child in si_el:
        if child.tag==q('t'):
            out.append(child.text or '')
        elif child.tag==q('r'):
            t=child.find(q('t'))
            out.append('' if t is None else (t.text or ''))
    return ''.join(out)

def _load_shared_strings(z,part):
    if part is None:return []
    strings=[]
    for _,el in ET.iterparse(io.BytesIO(z.read(part)),events=('end',)):
        if el.tag==q('si'):
            strings.append(_si_text(el))
            el.clear()
    return strings
def _load_styles(z,part):
    if part is None:return {'xfs':[],'custom_date':set(),'custom_percent':set()}
    root=ET.fromstring(z.read(part))
    custom={}
    nf=root.find(q('numFmts'))
    if nf is not None:
        for f in nf.findall(q('numFmt')):
            custom[int(f.get('numFmtId'))]=(f.get('formatCode') or '').lower()
    xfs=[]
    cellxfs=root.find(q('cellXfs'))
    if cellxfs is not None:
        for xf in cellxfs.findall(q('xf')):
            xfs.append(int(xf.get('numFmtId',0)))
    custom_date=set();custom_percent=set()
    for fid,code in custom.items():
        if '%' in code:custom_percent.add(fid)
        if any(ch in code for ch in ('y','d','h','s')) or ('m' in code and 'red' not in code):
            custom_date.add(fid)
    return {'xfs':xfs,'custom_date':custom_date,'custom_percent':custom_percent}

def _style_kind(styles,xf_id,date_style_ids):
    """百分比只作样式分类说明，不改写底层Decimal；日期只认配置允许的样式。

    date_style_ids按资料约定指cellXfs样式索引（即单元格s属性），不是numFmtId。
    """
    xfs=styles['xfs']
    if not isinstance(xf_id,int) or not 0<=xf_id<len(xfs):return 'unknown',None
    if xf_id in date_style_ids:return 'date-configured',xfs[xf_id]
    fid=xfs[xf_id]
    if fid in BUILTIN_DATE_FMTS or fid in BUILTIN_TIME_FMTS or fid in styles['custom_date']:
        return 'date-unconfigured',fid
    if fid==10 or fid in styles['custom_percent']:return 'percent',fid
    return 'other',fid

def _convert_date(serial_text,date1904):
    """按workbook纪元转换；1900序号60（含小数）标为虚构日期而非合法公历日。"""
    d=Decimal(serial_text)
    import math
    days=math.floor(d)
    if not date1904:
        if days==60:return None,'fictional-1900-02-29'
        ordinal=EPOCH_1900_ORDINAL+days
    else:
        ordinal=EPOCH_1904_ORDINAL+days
    try:day=date.fromordinal(ordinal)
    except ValueError:return None,'serial-out-of-calendar-range'
    frac=d-days
    if frac:
        total=(frac*86400).to_integral_value()
        hh,rem=divmod(int(total),3600);mm,ss=divmod(rem,60)
        return '%sT%02d:%02d:%02d'%(day.isoformat(),hh,mm,ss),None
    return day.isoformat(),None

def _cell_snapshot(el):
    v=is_text=None;f=None;f_type=f_si=f_ref=None
    for child in el:
        if child.tag==q('f'):
            f=child.text if child.text is not None else ''
            f_type=child.get('t');f_si=child.get('si');f_ref=child.get('ref')
        elif child.tag==q('v'):
            v=child.text
        elif child.tag==q('is'):
            is_text=''.join((t.text or '') for t in child.iter(q('t')))
    return {'r':el.get('r'),'t':el.get('t'),'s':el.get('s'),
            'f':f,'f_type':f_type,'f_si':f_si,'f_ref':f_ref,'v':v,'is':is_text}

def _iterparse_sheet(z,sheet_part):
    """iterparse行节点，行处理完即从父节点摘除并clear，避免空元素在根上堆积。

    注意不能在f元素的end事件clear：单元格快照要等整行结束才取，提前clear会
    丢掉公式文本与t/si/ref属性。
    """
    stream=io.BytesIO(z.read(sheet_part))
    sheet_data=None
    for event,el in ET.iterparse(stream,events=('start','end')):
        if event=='start':
            if el.tag==q('sheetData'):sheet_data=el
            continue
        if el.tag==q('row'):
            row={'r':int(el.get('r')),'hidden':el.get('hidden')=='1',
                 'cells':[_cell_snapshot(c) for c in el if c.tag==q('c')]}
            if sheet_data is not None:sheet_data.remove(el)
            el.clear()
            yield row
        elif el.tag in (q('mergeCell'),q('col'),q('autoFilter'),q('dimension')):
            el.clear()

def _collect_index(z,sheet_part,hf,hl,strings):
    """第一遍流式索引：合并区、隐藏列、筛选、共享公式基格、表头快照、出现过的列。"""
    merges={};hidden_cols=set();autofilter=None;dimension=None
    merge_refs=[];masters={};conflicting=set();used_si=set()
    seen_cols=set();header_cells=[]
    current_cell=None
    stream=io.BytesIO(z.read(sheet_part))
    for event,el in ET.iterparse(stream,events=('start','end')):
        if event=='start':
            if el.tag==q('c'):current_cell=el.get('r')
            continue
        if el.tag==q('mergeCell'):
            ref=el.get('ref');merge_refs.append(ref)
            rng=ref.split(':')
            r1,c1=split_ref(rng[0])
            if len(rng)==1:
                merges[(r1,c1)]=rng[0];seen_cols.add(c1)
            else:
                r2,c2=split_ref(rng[1]);top=rng[0]
                for rr in range(r1,r2+1):
                    for cc in range(c1,c2+1):
                        merges[(rr,cc)]=top;seen_cols.add(cc)
        elif el.tag==q('col') and el.get('hidden')=='1':
            hidden_cols.update(range(int(el.get('min')),int(el.get('max'))+1))
        elif el.tag==q('autoFilter'):
            autofilter=el.get('ref')
        elif el.tag==q('dimension'):
            dimension=el.get('ref')
        elif el.tag==q('f') and el.get('t')=='shared':
            si=el.get('si');used_si.add(si)
            if el.get('ref'):
                master={'ref':el.get('ref'),'formula':el.text,'cell':current_cell}
                if si in masters and masters[si]!=master:conflicting.add(si)
                else:masters.setdefault(si,master)
        elif el.tag==q('c') and el.get('r'):
            r,c=split_ref(el.get('r'));seen_cols.add(c)
            if hf<=r<=hl:
                snap=_cell_snapshot(el)
                snap['header_text']=_header_text(snap,strings)[0]
                header_cells.append(snap)
            el.clear()
    return {'merges':merges,'merge_refs':merge_refs,'hidden_cols':hidden_cols,
            'autofilter':autofilter,'dimension':dimension,'masters':masters,
            'conflicting':conflicting,'used_si':used_si,'seen_cols':seen_cols,
            'header_cells':header_cells}

def _resolve_parts(z,sheet_name):
    names=_members(z)
    if '[Content_Types].xml' not in names:raise ExtractError('缺少内容类型成员')
    if '_rels/.rels' not in names:raise ExtractError('缺少包根关系成员')
    root_rels=_rels(z.read('_rels/.rels'))
    wb_rid=next((rid for rid,r in root_rels.items()
                 if r['type'].endswith('/officeDocument') and r['mode']!='External'),None)
    if wb_rid is None:raise ExtractError('缺少工作簿文档关系')
    wb_part=_resolve('',root_rels[wb_rid]['target'])
    if wb_part not in names:raise ExtractError('工作簿成员不存在:'+wb_part)
    wb=ET.fromstring(z.read(wb_part))
    wbpr=wb.find(q('workbookPr'))
    date1904=wbpr is not None and wbpr.get('date1904') in ('1','true')
    wanted=next((sh for sh in wb.find(q('sheets')) if sh.get('name')==sheet_name),None)
    if wanted is None:raise ExtractError('配置的工作表名称不存在:'+sheet_name)
    rid=wanted.get('{%s}id'%RNS)
    rels_part=_rels_part_of(wb_part)
    rels=_rels(z.read(rels_part)) if rels_part in names else {}
    if rid not in rels:raise ExtractError('工作表关系ID缺失:'+rid)
    rel=rels[rid]
    if rel['mode']=='External':raise ExtractError('工作表关系是外部目标，不抓取外部工作簿')
    if not rel['type'].endswith('/worksheet'):raise ExtractError('关系类型不是工作表:'+rid)
    sheet_part=_resolve(wb_part,rel['target'])
    if sheet_part not in names:raise ExtractError('工作表成员不存在:'+sheet_part)
    ss_part=style_part=None;external=[]
    for rid2,r in rels.items():
        if r['mode']=='External':
            external.append({'rid':rid2,'target':r['target'],'type':r['type']});continue
        if r['type'].endswith('/sharedStrings'):ss_part=_resolve(wb_part,r['target'])
        elif r['type'].endswith('/styles'):style_part=_resolve(wb_part,r['target'])
    return {'workbook_part':wb_part,'sheet_part':sheet_part,'sheet_rid':rid,
            'shared_strings_part':ss_part,'styles_part':style_part,
            'date1904':date1904,'external_rels':external}

def _header_text(cell,strings):
    if cell['t']=='s' and cell['v'] is not None:
        try:return strings[int(cell['v'])],'shared-string'
        except (IndexError,ValueError):return cell['v'],'shared-index-invalid'
    if cell['t']=='inlineStr':return cell['is'] or '','inline-string'
    if cell['t']=='b':return ('TRUE' if cell['v']=='1' else 'FALSE'),'boolean'
    return cell['v'] or '','plain'

def _header_paths(idx,strings,total_cols,hf,hl):
    """逐列有序表头路径。

    合并区只取左上角作为表头来源；落在合并覆盖区的列不产生自己的条目，
    而是在组装路径时继承左上角单元格的条目（不能把合并值写成正文每格真实值）。
    """
    columns={}
    for cell in idx['header_cells']:
        r,c=split_ref(cell['r'])
        if not (hf<=r<=hl):continue
        top=idx['merges'].get((r,c))
        if top is not None and top!=cell['r']:
            continue
        _,kind=_header_text(cell,strings)
        columns.setdefault(c,[]).append(
            {'value':cell.get('header_text',''),'source':cell['r'],'origin_type':kind})
    # 合并左上角按列建立继承条目
    inherited={}
    for (r,c),topleft in idx['merges'].items():
        if not (hf<=r<=hl):continue
        tr,tc=split_ref(topleft)
        if tr==r and tc==c:continue
        for cell in idx['header_cells']:
            if cell['r']==topleft:
                _,kind=_header_text(cell,strings)
                inherited.setdefault(c,[]).append(
                    {'value':cell.get('header_text',''),'source':topleft,'origin_type':kind})
    paths={}
    for c in range(1,total_cols+1):
        own=columns.get(c,[])
        extra=[e for e in inherited.get(c,[]) if e['source'] not in {o['source'] for o in own}]
        # 继承项来自更靠上的合并行，排在本列自有条目之前并按来源行排序
        paths[c]=sorted(own+extra,key=lambda e:split_ref(e['source'])[0])
    return paths

def _decimal(raw):
    try:return Decimal(raw)
    except (InvalidOperation,ValueError):return None

def _formula_info(ctx,cell,coord):
    info={'raw':cell['f'],'expanded':cell['f'],'expand_reason':None,
          'kind':cell['f_type'] or 'normal','si':cell['f_si'],
          'master_ref':cell['f_ref'],'master_cell':None}
    if cell['f_type']!='shared':
        reasons=unsupported_reasons(cell['f'] or '')
        if reasons:info['expand_reason']=';'.join(reasons)
        return info
    index=ctx['index'];si=cell['f_si']
    if si in index['conflicting']:
        info['expand_reason']='共享公式组存在多个冲突基格';return info
    master=index['masters'].get(si)
    if master is None or master['formula'] is None:
        info['expand_reason']='共享公式组缺少基格';return info
    info['master_cell']=master['cell']
    if cell['f_ref']:
        info['expanded']=master['formula'];return info
    expanded,reason=translate_shared(master['formula'],master['cell'],coord)
    info['expanded']=expanded;info['expand_reason']=reason
    return info

def _cell_record(ctx,row_index,row_hidden,col,cell):
    index=ctx['index']
    rec={'package_member':ctx['sheet_part'],'sheet_rid':ctx['sheet_rid'],
         'coordinate':join_ref(row_index,col),'header_path':ctx['header_paths'].get(col,[]),
         'row_hidden':row_hidden,'col_hidden':col in index['hidden_cols'],
         'style':None,'formula':None,'cache_state':'not-a-formula'}
    xf=None
    if cell is not None and cell['s'] is not None:
        try:xf=int(cell['s'])
        except ValueError:xf=None
    elif cell is not None and ctx['styles']['xfs']:
        xf=0
    kind,fmt=_style_kind(ctx['styles'],xf,ctx['date_style_ids'])
    rec['style']={'xf_id':xf,'raw_s':None if cell is None else cell['s'],
                  'num_fmt_id':fmt,'classification':kind}
    if cell is None:
        rec.update(origin_type='missing',value=None,state='missing-cell')
        return rec
    rec['origin_type']=cell['t'] or 'n'
    if cell['f'] is not None:
        rec['formula']=_formula_info(ctx,cell,rec['coordinate'])
    f_present=rec['formula'] is not None
    t=cell['t']
    if f_present:
        if cell['v'] is None:
            rec.update(value=None,state='formula-no-cache',cache_state='no-cache');return rec
        rec['cache_state']='cached'
    if t=='s':
        if cell['v'] is None:
            rec.update(value=None,state='shared-index-missing')
        else:
            try:rec.update(value=ctx['strings'][int(cell['v'])],state='shared-string')
            except (IndexError,ValueError):rec.update(value=cell['v'],state='shared-index-invalid')
    elif t=='inlineStr':
        if cell['is'] is None or cell['is']=='':
            rec.update(value='',state='empty-inline-string')
        else:rec.update(value=cell['is'],state='inline-string')
    elif t=='b':
        if cell['v'] in ('0','1'):rec.update(value=(cell['v']=='1'),state='boolean')
        else:rec.update(value=cell['v'],state='boolean-invalid')
    elif t=='e':
        rec.update(value=cell['v'],state='error')
    elif t=='str':
        rec.update(value=cell['v'],state='formula-string' if f_present else 'string')
    elif t in (None,'n'):
        if cell['v'] is None:
            rec.update(value=None,state='empty-number')
        else:
            dec=_decimal(cell['v'])
            entry={'lexical':cell['v'],
                   'decimal':None if dec is None else str(dec),'percent':kind=='percent'}
            if kind=='date-configured':
                iso,note=_convert_date(cell['v'],ctx['date1904'])
                entry['date_iso']=iso;entry['date_note']=note
            rec.update(value=entry,state='number')
    else:
        rec.update(value=cell['v'],state='unknown-type-%s'%t)
    return rec

def _validate(config):
    if not isinstance(config.get('sheet'),str) or not config['sheet']:
        raise ExtractError('sheet必须为非空名称')
    hf=integer(config['header_first'],1);hl=integer(config['header_last'],1)
    bf=integer(config['body_first'],1);bl=integer(config['body_last'],1)
    if hl<hf:raise ExtractError('表头行逆序')
    if bl<bf:raise ExtractError('正文行逆序')
    limit=integer(config.get('row_cache_limit',64),1)
    ds=config.get('date_style_ids',[])
    if not isinstance(ds,list) or any(type(x) is not int or x<0 for x in ds):
        raise ExtractError('date_style_ids必须为非负整数列表')
    return hf,hl,bf,bl,limit,set(ds)

class WorkbookExtract:
    """保持ZIP打开的流式采集句柄；records()为有界行缓存生成器。"""
    def __init__(self,raw,config):
        self._cfg=_validate(config)
        self._zip=zipfile.ZipFile(io.BytesIO(raw))
        try:
            hf,hl,bf,bl,limit,ds=self._cfg
            parts=_resolve_parts(self._zip,config['sheet'])
            self._strings=_load_shared_strings(self._zip,parts['shared_strings_part'])
            self._styles=_load_styles(self._zip,parts['styles_part'])
            self._index=_collect_index(self._zip,parts['sheet_part'],hf,hl,self._strings)
            total=max(self._index['seen_cols'],default=0)
            self._total_cols=total
            self._header_paths=_header_paths(self._index,self._strings,total,hf,hl)
            diagnostics=self._diagnostics()
            self.meta={**parts,'dimension':self._index['dimension'],
                       'autofilter':self._index['autofilter'],
                       'merged_areas':sorted(self._index['merge_refs']),
                       'hidden_columns':sorted(self._index['hidden_cols']),
                       'header_rows':[hf,hl],'body_rows':[bf,bl],
                       'total_columns':total,'row_cache_limit':limit,
                       'diagnostics':diagnostics}
        except BaseException:
            self._zip.close();raise

    def _diagnostics(self):
        out=[];idx=self._index
        for si in sorted(idx['used_si']):
            if si in idx['conflicting']:
                out.append({'code':'shared-formula-conflicting-masters','si':si})
            elif si not in idx['masters']:
                out.append({'code':'shared-formula-missing-master','si':si})
            else:
                r1,c1,r2,c2=split_range(idx['masters'][si]['ref'])
                if idx['masters'][si]['cell']!=join_ref(r1,c1):
                    out.append({'code':'shared-formula-master-not-at-ref-origin','si':si})
        return out

    def records(self):
        hf,hl,bf,bl,limit,_=self._cfg
        ctx={'sheet_part':self.meta['sheet_part'],'sheet_rid':self.meta['sheet_rid'],
             'strings':self._strings,'styles':self._styles,'index':self._index,
             'header_paths':self._header_paths,'date1904':self.meta['date1904'],
             'date_style_ids':self._cfg[5]}
        heap=[];tie=0
        def emit(row):
            present={split_ref(c['r'])[1]:c for c in row['cells']}
            for col in range(1,self._total_cols+1):
                yield _cell_record(ctx,row['r'],row['hidden'],col,present.get(col))
        for row in _iterparse_sheet(self._zip,self.meta['sheet_part']):
            if not (bf<=row['r']<=bl):continue
            heapq.heappush(heap,(row['r'],tie,row));tie+=1
            if len(heap)>=limit:
                yield from emit(heapq.heappop(heap)[2])
        while heap:
            yield from emit(heapq.heappop(heap)[2])

    def close(self):self._zip.close()
    def __enter__(self):return self
    def __exit__(self,*exc):self.close()

def open_workbook(raw,config):
    """返回流式采集句柄；调用方负责close或使用with。"""
    return WorkbookExtract(raw,config)

def extract(raw,config):
    """便利接口：返回(meta, 全部记录)。大工作表请改用open_workbook流式接口。"""
    with open_workbook(raw,config) as wb:
        return wb.meta,list(wb.records())
