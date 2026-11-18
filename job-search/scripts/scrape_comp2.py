import re,json,time,subprocess,sys,threading
from concurrent.futures import ThreadPoolExecutor
slugs=json.load(open('comp_slugs.json'))[150:]
res=json.load(open('comp_taif.json')); lock=threading.Lock(); n=[0]
def get(u):
    return subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',u],capture_output=True).stdout.decode('utf-8','ignore')
def work(sl):
    h=get("https://sabbar.com/en/jobs/companies/"+sl); time.sleep(0.5)
    name='';desc='';jobs=[]
    for m in re.finditer(r'<script type="application/ld\+json">(.*?)</script>',h,re.S):
        try: d=json.loads(m.group(1))
        except: continue
        if d.get('@type')=='Organization': name=d.get('name','');desc=d.get('description','')
        if d.get('@type')=='ItemList': jobs+=[(j.get('url',''),j.get('name','')) for j in d.get('itemListElement',[])]
    taif=[j for j in jobs if '/c-taif' in j[0]]
    with lock:
        if taif: res[sl]=dict(name=name,desc=desc[:500],taif_jobs=taif,total_jobs=len(jobs))
        n[0]+=1
        if n[0]%100==0:
            print(n[0],len(res),flush=True); json.dump(res,open('comp_taif.json','w'),ensure_ascii=False)
with ThreadPoolExecutor(3) as ex: list(ex.map(work,slugs))
json.dump(res,open('comp_taif.json','w'),ensure_ascii=False)
print('done comp',len(res),flush=True)
