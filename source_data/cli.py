"""显示资料库存与原始字段摘要，并提供xlsx逐格来源化采集。"""
import argparse
import json
from pathlib import Path
from .io import load,asset
from .domain import summarize
from .xlsx import open_workbook

def main():
    p=argparse.ArgumentParser()
    sub=p.add_subparsers(dest='command',required=True)
    ins=sub.add_parser('inspect');ins.add_argument('file')
    col=sub.add_parser('collect');col.add_argument('file');col.add_argument('out_jsonl')
    a=p.parse_args()
    value,digest=load(a.file)
    root=Path(a.file).parent
    if a.command=='inspect':
        print(json.dumps({'source_sha256':digest,
                          'summary':summarize(value,root)},ensure_ascii=False,indent=2))
        return
    raw,wb_digest=asset(root,value['workbook'])
    with open_workbook(raw,value) as wb,open(a.out_jsonl,'w',encoding='utf8') as out:
        out.write(json.dumps({'meta':{**wb.meta,'workbook_sha256':wb_digest}},
                             ensure_ascii=False)+'\n')
        n=0
        for rec in wb.records():
            out.write(json.dumps(rec,ensure_ascii=False)+'\n');n+=1
    print('已写出%d条逐格记录 -> %s'%(n,a.out_jsonl))
if __name__=='__main__':main()
