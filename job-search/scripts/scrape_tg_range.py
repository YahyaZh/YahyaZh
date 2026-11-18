import re,json,time,subprocess,sys,html as H
ch,start,stop=sys.argv[1],int(sys.argv[2]),int(sys.argv[3]); out={}; before=start; stall=0
fn='tg_%s_%d.json'%(ch,start)
while before>stop:
    h=subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',"https://t.me/s/%s?before=%d"%(ch,before)],capture_output=True).stdout.decode('utf-8','ignore')
    ids=[]
    for b in re.split(r'<div class="tgme_widget_message_wrap',h)[1:]:
        p=re.search(r'data-post="[^/]+/(\d+)"',b)
        if not p: continue
        pid=int(p.group(1)); ids.append(pid)
        t=re.search(r'<div class="tgme_widget_message_text[^"]*"[^>]*>(.*?)</div>',b,re.S)
        d=re.search(r'<time datetime="([^"]+)"',b)
        txt=re.sub(r'<[^>]+>','',H.unescape(re.sub(r'<br\s*/?>','\n',t.group(1)))) if t else ''
        out[pid]=dict(id=pid,date=d.group(1) if d else '',text=txt.strip(),links=[l for l in re.findall(r'href="(https?://[^"]+)"',b) if 't.me' not in l][:5])
    if not ids:
        stall+=1; before-=20
        if stall>20: break
        continue
    stall=0
    before=min(ids)
    if len(out)%500<20: json.dump(out,open(fn,'w'),ensure_ascii=False)
    time.sleep(0.5)
json.dump(out,open(fn,'w'),ensure_ascii=False)
print('done tg range',start,len(out),flush=True)
