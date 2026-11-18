import re,json,time,subprocess,sys,html as H
base="https://www.wdeftksa.com/sa/jobs/city-%D9%88%D8%B8%D8%A7%D8%A6%D9%81-%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81"
out=[]
card=re.compile(r'<li><a href="(/sa/jobs/\d+)" class="wd-jobcard">(.*?)</a></li>',re.S)
for p in range(1,560):
    u=base+("?page=%d"%p if p>1 else "")
    r=subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',u],capture_output=True)
    h=r.stdout.decode('utf-8','ignore')
    cs=card.findall(h)
    if not cs: print('empty',p,file=sys.stderr); 
    for href,body in cs:
        t=re.search(r'wd-jobcard__title">(.*?)</h3>',body,re.S)
        co=re.search(r'meta-strong">.*?<span>(.*?)</span>',body,re.S)
        loc=re.search(r'meta-location" title="(.*?)"',body,re.S)
        dt=re.search(r'<time datetime="(.*?)"',body)
        tag=re.findall(r'class="wd-tag">(.*?)</span>',body)
        out.append(dict(url="https://www.wdeftksa.com"+href,title=H.unescape(t.group(1)) if t else '',company=H.unescape(co.group(1)) if co else '',loc=H.unescape(loc.group(1)) if loc else '',date=dt.group(1) if dt else '',tags=tag,page=p))
    if p%25==0:
        json.dump(out,open('wdeft_items.json','w'),ensure_ascii=False); print(p,len(out),file=sys.stderr)
    time.sleep(0.6)
json.dump(out,open('wdeft_items.json','w'),ensure_ascii=False)
print('done',len(out))
