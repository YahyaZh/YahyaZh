import re,json,time,subprocess,sys
def get(u):
    return subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',u],capture_output=True).stdout.decode('utf-8','ignore')
slugs=[]
for p in range(1,250):
    h=get("https://sabbar.com/en/jobs/companies?page=%d"%p)
    s=sorted(set(re.findall(r'href="/en/jobs/companies/([^"?/]+)"',h)))
    if not s: break
    for x in s:
        if x not in slugs: slugs.append(x)
    time.sleep(0.5)
print('slugs',len(slugs),file=sys.stderr)
json.dump(slugs,open('comp_slugs.json','w'))
res={}
for i,sl in enumerate(slugs):
    h=get("https://sabbar.com/en/jobs/companies/"+sl)
    name='';desc='';jobs=[]
    for m in re.finditer(r'<script type="application/ld\+json">(.*?)</script>',h,re.S):
        try: d=json.loads(m.group(1))
        except: continue
        if d.get('@type')=='Organization': name=d.get('name','');desc=d.get('description','')
        if d.get('@type')=='ItemList': jobs+=[(j.get('url',''),j.get('name','')) for j in d.get('itemListElement',[])]
    taif=[j for j in jobs if '/c-taif' in j[0]]
    if taif: res[sl]=dict(name=name,desc=desc[:500],taif_jobs=taif,total_jobs=len(jobs))
    if i%50==0:
        print(i,len(res),file=sys.stderr); json.dump(res,open('comp_taif.json','w'),ensure_ascii=False)
    time.sleep(0.5)
json.dump(res,open('comp_taif.json','w'),ensure_ascii=False)
print('done',len(slugs),len(res))
