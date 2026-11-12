"""公式词法与共享公式引用平移。不执行公式，也不接触外部工作簿。

支持词法见docs/data-contract.md：A1单格/矩形范围、可带单引号转义表名、
绝对/相对/混合坐标、数字、双引号字符串、算术比较运算、括号、逗号和普通函数名。
结构化表引用、外部工作簿引用、动态数组语法与R1C1一律不展开，只保留原式与原因。
"""
import re
from .coords import col_to_index,col_to_letters,split_ref,MAX_ROW,MAX_COL

_SHEET_NAME=r"(?:'(?:''|[^'])*'|[A-Za-z_一-鿿][A-Za-z0-9_.一-鿿]*)"
_CELL=r"\$?[A-Z]{1,3}\$?\d{1,7}"
_REF_RE=re.compile(
    r"(?:%s\s*!)?%s(?::%s)?(?![A-Za-z0-9_])"%(_SHEET_NAME,_CELL,_CELL),re.ASCII)
_R1C1_RE=re.compile(r'(?<![A-Za-z_.$])(?:R|C)(?:\$?\d+|\[-?\d+\])')
_STRUCT_RE=re.compile(r"(?:[A-Za-z_一-鿿][\w一-鿿]*\.)?\[[^\]]+\]")

def _split_code(formula):
    """切出字符串字面量区间与代码区间，字符串内的内容永远不按引用解释。"""
    parts=[];i=0;start=0;n=len(formula)
    while i<n:
        if formula[i]=='"':
            if start<i:parts.append(('code',formula[start:i]))
            j=i+1
            while j<n:
                if formula[j]=='"':
                    if j+1<n and formula[j+1]=='"':j+=2;continue
                    break
                j+=1
            if j>=n:
                parts.append(('string',formula[i:]))
                i=n;start=n;break
            parts.append(('string',formula[i:j+1]))
            i=j+1;start=i
        else:i+=1
    if start<n:parts.append(('code',formula[start:]))
    return parts

def unsupported_reasons(formula):
    """返回该公式中超出支持词法的语法原因列表（空列表表示可解析）。"""
    reasons=[]
    for kind,seg in _split_code(formula):
        if kind=='string':continue
        if '[' in seg and ']' in seg:
            reasons.append('结构化表引用或外部工作簿引用方括号语法')
        if '[' in seg and ']' not in seg or ']' in seg and '[' not in seg:
            reasons.append('不配对的方括号语法')
        if _R1C1_RE.search(seg):reasons.append('R1C1引用样式')
        if _STRUCT_RE.search(seg):reasons.append('结构化表引用')
        if '@' in seg:reasons.append('隐式交叉/动态数组专用语法')
    return reasons

def _translate_cell(tok,dr,dc):
    dollar_col,letters,dollar_row,digits=tok
    col=col_to_index(letters);row=int(digits)
    if not dollar_col:
        col+=dc
        if not 1<=col<=MAX_COL:return None,'列平移越界'
    if not dollar_row:
        row+=dr
        if not 1<=row<=MAX_ROW:return None,'行平移越界'
    return ('$' if dollar_col else '')+col_to_letters(col)+('$' if dollar_row else '')+str(row),None

_CELL_TOKEN_RE=re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d{1,7})')

def translate_shared(formula,master_ref,slave_ref):
    """把基公式按基格->从属格的位移展开，返回(展开式,原因)。

    原因非None时不保证展开式可用，调用方应同时保留原式。
    带表名（单引号或裸表名）的引用不做平移；双引号字符串原样保留。
    """
    reasons=unsupported_reasons(formula)
    if reasons:return formula,';'.join(reasons)
    mr,mc=split_ref(master_ref);sr,sc=split_ref(slave_ref)
    dr=sr-mr;dc=sc-mc
    out=[]
    for kind,seg in _split_code(formula):
        if kind=='string':
            out.append(seg);continue
        pos=0
        for m in _REF_RE.finditer(seg):
            out.append(seg[pos:m.start()])
            text=m.group(0)
            bang=text.find('!')
            if bang>=0:
                out.append(text)
            else:
                cells=text.split(':')
                rendered=[]
                bad=None
                for cell in cells:
                    cm=_CELL_TOKEN_RE.fullmatch(cell)
                    new,err=_translate_cell(cm.groups(),dr,dc)
                    if err:bad=err;break
                    rendered.append(new)
                if bad:return formula,'相对引用平移超出工作表边界:'+bad
                out.append(':'.join(rendered))
            pos=m.end()
        out.append(seg[pos:])
    return ''.join(out),None
