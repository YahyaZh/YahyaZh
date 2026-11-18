import re,json,time,subprocess,sys,html as H
from concurrent.futures import ThreadPoolExecutor
def get(u): return subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',u],capture_output=True).stdout.decode('utf-8','ignore')
ids=[];empty=0
for p in range(1,1000):
    h=get("https://www.ewdifh.com/search?word=%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81&page="+str(p))
    new=[i for i in re.findall(r'href="https://www\.ewdifh\.com/jobs/(\d+)"',h) if i not in ids]
    ids+=list(dict.fromkeys(new))
    if not new:
        empty+=1
        if empty>=2: break
    else: empty=0
    if p%20==0: print('search',p,len(ids),flush=True)
    time.sleep(0.4)
print('ids',len(ids),flush=True)
out={}
def det(i):
    h=get("https://www.ewdifh.com/jobs/"+i); time.sleep(0.4)
    t=re.sub(r'<script.*?</script>|<style.*?</style>','',h,flags=re.S)
    title=re.search(r'<h1[^>]*>(.*?)</h1>',t,re.S)
    body=H.unescape(re.sub(r'\s+',' ',re.sub(r'<[^>]+>',' ',t)))
    s=body.find('بواسطة'); 
    d=re.search(r'(\d{2}-\d{2}-\d{4})',body)
    out[i]=dict(id=i,title=H.unescape(re.sub(r'<[^>]+>','',title.group(1))).strip() if title else '',date=d.group(1) if d else '',text=body[s:s+2500] if s>0 else body[:2500])
with ThreadPoolExecutor(3) as ex:
    for n,_ in enumerate(ex.map(det,ids)):
        if n%200==0: print('det',n,flush=True); json.dump(out,open('ew_items.json','w'),ensure_ascii=False)
json.dump(out,open('ew_items.json','w'),ensure_ascii=False)
print('done ew',len(out),flush=True)
