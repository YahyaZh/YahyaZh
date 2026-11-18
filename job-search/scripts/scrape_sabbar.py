import re,json,time,subprocess,urllib.parse,sys
urls=[u.strip() for u in open('raw/ctaif_urls.txt') if '/ar/jobs/' in u]
seen={}
def get(u):
    r=subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',u],capture_output=True)
    return r.stdout.decode('utf-8','ignore')
def items(h):
    out=[]
    for m in re.finditer(r'<script type="application/ld\+json">(.*?)</script>',h,re.S):
        try: d=json.loads(m.group(1))
        except: continue
        if d.get('@type')=='ItemList':
            for i in d.get('itemListElement',[]): out.append(i)
    return out
q=list(urls); done=set()
while q:
    u=q.pop(0)
    if u in done: continue
    done.add(u)
    h=get(u); time.sleep(0.7)
    its=items(h)
    for i in its:
        seen.setdefault(i['url'],{'name':i.get('name',''),'desc':i.get('description','')[:600],'from':u})
    # pagination
    if '?page=' not in u and len(its)>=18:
        mx=max([int(x) for x in re.findall(r'[?&]page=(\d+)',h)]+[1])
        for p in range(2,min(mx,40)+1): q.append(u+'?page=%d'%p)
    print(len(done),len(q),len(seen),file=sys.stderr)
json.dump(seen,open('sabbar_taif_items.json','w'),ensure_ascii=False,indent=0)
print('items',len(seen))
