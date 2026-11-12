"""列出工作簿的原始包成员，不解释工作表或公式。"""
import io
import zipfile
from .io import asset,integer

def summarize(v,root):
    raw,digest=asset(root,v['workbook'])
    integer(v['header_first'],1);integer(v['header_last'],1)
    if v['header_last']<v['header_first']:raise ValueError('表头行逆序')
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        names=z.namelist()
        if len(set(names))!=len(names):raise ValueError('工作簿包成员重名')
        if '[Content_Types].xml' not in names:raise ValueError('缺少内容类型成员')
        return {'members':[{'name':i.filename,'bytes':i.file_size} for i in z.infolist()],'archive_sha256':digest,'requested_sheet':v['sheet']}
