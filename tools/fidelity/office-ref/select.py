import json,zipfile,sys,re,random,os
B=os.environ["EO_CORPUS_PUBLIC"]
fmt=sys.argv[1]; N=40
rows=[json.loads(l) for l in open(B+"batches/manifest.jsonl")]
rows=[r for r in rows if r['format']==fmt]
info=[]
for r in rows:
    p=B+r['file']; sz=os.path.getsize(p)
    d=dict(id=r['sha256'],input=p,bytes=sz,upstream=r['provenance'].get('upstreamPath','')[-80:])
    if sz>20e6: d['skip']='>20MB'; info.append(d); continue
    try: z=zipfile.ZipFile(p); names=z.namelist()
    except Exception as e: d['skip']='not a zip (encrypted/corrupt)'; info.append(d); continue
    if any('vbaProject' in n for n in names): d['skip']='macros'; info.append(d); continue
    pre='xl/' if fmt=='xlsx' else 'ppt/'
    unit='xl/worksheets/sheet' if fmt=='xlsx' else 'ppt/slides/slide'
    d['units']=len([n for n in names if n.startswith(unit) and n.endswith('.xml')])
    d['charts']=len([n for n in names if re.match(pre+'charts/chart[0-9]*\.xml',n)])
    d['media']=len([n for n in names if n.startswith(pre+'media/')])
    d['tables']=len([n for n in names if n.startswith('xl/tables/')]) if fmt=='xlsx' else 0
    d['drawings']=len([n for n in names if n.startswith(pre+'drawings/drawing')])
    try:
        wb=z.read('xl/workbook.xml' if fmt=='xlsx' else 'ppt/presentation.xml').decode('utf8','replace')
        if 'workbookProtection' in wb and 'workbookPassword' in wb: d['skip']='workbook password'
        if fmt=='xlsx':
            d['printareas']=wb.count('_xlnm.Print_Area'); d['hidden_sheets']=len(re.findall(r'<sheet [^>]*state="(?:hidden|veryHidden)"',wb))
        if fmt=='pptx': d['ext_links']=len([n for n in names if 'externalLinks' in n or 'embeddings' in n])
    except Exception: pass
    if any(n.startswith(('xl/externalLinks',)) for n in names): d['extlinks']=True
    d['score']=min(d['units'],6)*1+d['charts']*3+min(d['media'],5)*2+d['tables']*2+(1 if d['bytes']>100000 else 0)+ (2 if d['drawings'] else 0)
    info.append(d)
ok=[d for d in info if 'skip' not in d and d['units']>0]
json.dump(info,open(sys.argv[2],'w'))
random.seed(341)
# 25 top by score (feature-rich), 15 stratified rest (size/varied)
ok.sort(key=lambda d:(-d['score'],d['id']))
top=ok[:26]; rest=ok[26:]
rest.sort(key=lambda d:d['bytes'])
step=max(1,len(rest)//14); extra=rest[::step][:14]
sel=top+extra
seen=set(); out=[]
for d in sel:
    if d['id'] in seen: continue
    seen.add(d['id']); out.append(d)
out=out[:N]
print(fmt,'eligible',len(ok),'selected',len(out),'skipped',[ (d['skip']) for d in info if 'skip' in d][:10])
with open(sys.argv[3],'w') as f:
    for d in out:
        why=[]
        if d['units']>1: why.append('%d %s'%(d['units'],'sheets' if fmt=='xlsx' else 'slides'))
        if d['charts']: why.append('%d charts'%d['charts'])
        if d['media']: why.append('%d pictures'%d['media'])
        if d['tables']: why.append('%d tables'%d['tables'])
        if d['drawings']: why.append('drawings/shapes')
        if d.get('printareas'): why.append('print areas')
        why.append('size %d B'%d['bytes'])
        d2=dict(d); d2.pop('score'); d2['format']=fmt; d2['reason']='; '.join(why)+(' (feature-rich pick)' if d in top else ' (size-stratified pick)')
        f.write(json.dumps(d2)+'\n')
