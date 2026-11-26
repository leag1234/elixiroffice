import json,sys,os
kind,sel,refdir,out=sys.argv[1:5]
n=0; bad=[]
with open(out,'w') as f:
    for l in open(sel):
        d=json.loads(l); pdf=f"{refdir}/{d['id']}.pdf"
        meta=json.load(open(f"{refdir}/{d['id']}.json",encoding='utf-8-sig')) if os.path.exists(f"{refdir}/{d['id']}.json") else {}
        if not os.path.exists(pdf): bad.append((d['id'][:10],d['upstream'][-40:],meta.get('error'))); continue
        f.write(json.dumps({"id":d['id'][:40],"input":d['input'],"reference":pdf,"reference_source":"ms-office-desktop-16.0.%s"%meta.get('excel_build',meta.get('ppt_build')),"format":kind})+"\n"); n+=1
print(kind,n,'pairs; no reference:',bad)
