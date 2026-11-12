"""A1坐标与稀疏单元格定位，按OOXML行列上限做边界判断。"""
import re

MAX_ROW=1048576
MAX_COL=16384
_CELL_RE=re.compile(r'^([A-Z]{1,3})(\d{1,7})$')

def col_to_index(letters):
    n=0
    for ch in letters:
        n=n*26+(ord(ch)-64)
    if not 1<=n<=MAX_COL:raise ValueError('列坐标越界')
    return n

def col_to_letters(index):
    if not 1<=index<=MAX_COL:raise ValueError('列序号越界')
    out=''
    while index:
        index,r=divmod(index-1,26)
        out=chr(65+r)+out
    return out

def split_ref(ref):
    m=_CELL_RE.match(ref)
    if not m:raise ValueError('非法单元格坐标:%r'%ref)
    col=col_to_index(m.group(1));row=int(m.group(2))
    if not 1<=row<=MAX_ROW:raise ValueError('行坐标越界')
    return row,col

def join_ref(row,col):
    return col_to_letters(col)+str(row)

def split_range(text):
    first,sep,last=text.partition(':')
    r1,c1=split_ref(first)
    if not sep:return r1,c1,r1,c1
    r2,c2=split_ref(last)
    return r1,c1,r2,c2
