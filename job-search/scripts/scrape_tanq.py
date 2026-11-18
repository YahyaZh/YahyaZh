import re,json,time,subprocess,sys,html as H
U="https://saudi.tanqeeb.com/ar/s/%D9%88%D8%B8%D8%A7%D8%A6%D9%81/%D9%88%D8%B8%D8%A7%D8%A6%D9%81-%D9%81%D9%8A-%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81"
out={};empty=0
for p in range(1,400):
    h=subprocess.run(['curl','-sS','-m','40','-A','Mozilla/5.0',U+("?page_no=%d"%p)],capture_output=True).stdout.decode('utf-8','ignore')
    arts=re.findall(r'<article id="JOB-[^"]*"(.*?)>',h,re.S)
    new=0
    for a in arts:
        attrs=dict((k,H.unescape(v)) for k,v in re.findall(r'data-job-([a-z-]+)="([^"]*)"',a))
        if attrs.get('id') and attrs['id'] not in out:
            out[attrs['id']]=attrs; new+=1
    print(p,len(arts),new,len(out),file=sys.stderr)
    if new==0:
        empty+=1
        if empty>=3: break
    else: empty=0
    if p%20==0: json.dump(out,open('tanq_items.json','w'),ensure_ascii=False)
    time.sleep(0.8)
json.dump(out,open('tanq_items.json','w'),ensure_ascii=False)
print('done',len(out))
