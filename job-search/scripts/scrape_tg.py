import re,json,time,subprocess,sys,html as H
ch=sys.argv[1]; out={}
before=None; stall=0
while True:
    u="https://t.me/s/%s"%ch+("?before=%d"%before if before else "")
    h=subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',u],capture_output=True).stdout.decode('utf-8','ignore')
    blocks=re.split(r'<div class="tgme_widget_message_wrap',h)[1:]
    ids=[]
    for b in blocks:
        p=re.search(r'data-post="[^/]+/(\d+)"',b)
        if not p: continue
        pid=int(p.group(1)); ids.append(pid)
        t=re.search(r'<div class="tgme_widget_message_text[^"]*"[^>]*>(.*?)</div>',b,re.S)
        d=re.search(r'<time datetime="([^"]+)"',b)
        links=re.findall(r'href="(https?://[^"]+)"',b)
        txt=H.unescape(re.sub(r'<br\s*/?>','\n',t.group(1))) if t else ''
        txt=re.sub(r'<[^>]+>','',txt)
        out[pid]=dict(id=pid,date=d.group(1) if d else '',text=txt.strip(),links=[l for l in links if 't.me' not in l][:5])
    if not ids: 
        stall+=1
        if stall>3: break
        time.sleep(3); continue
    stall=0
    mn=min(ids)
    if before is not None and mn>=before: break
    before=mn
    if len(out)%400<20: print(len(out),before,file=sys.stderr); json.dump(out,open('tg_%s.json'%ch,'w'),ensure_ascii=False)
    if before<=1: break
    time.sleep(0.8)
json.dump(out,open('tg_%s.json'%ch,'w'),ensure_ascii=False)
print('done',ch,len(out))
