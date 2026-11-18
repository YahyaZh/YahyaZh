import json,glob,re,subprocess,time,html as H,threading,os
from concurrent.futures import ThreadPoolExecutor
tg={}
for f in glob.glob('tg_jobstaif*.json'): tg.update(json.load(open(f)))
OTHER=re.compile(r'الرياض|جدة|جده|مكة|مكه|المدينة المنورة|الدمام|الخبر|القصيم|بريدة|أبها|ابها|خميس|تبوك|جازان|حائل|الأحساء|الاحساء|نجران|الجبيل|ينبع|عدة مدن|جميع مناطق|كافة مناطق|مناطق المملكة|عن بعد|عن بُعد|#عن_بعد|remote',re.I)
NONJOB=re.compile(r'دورة|دورات|ورشة|ندوة|استبيان|القبول|whatsapp|واتساب|pinned|ابتعاث|نتائج|تطوع')
todo={}
for v in tg.values():
    t=v['text']
    if not t or NONJOB.search(t[:300]): continue
    if 'الطائف' not in t and OTHER.search(t): continue
    m=re.search(r'saudijobs24\.com/(t\d+)',t)
    if m: todo[m.group(1)]=v['id']
out=json.load(open('sj_items.json')) if os.path.exists('sj_items.json') else {}
todo={k:v for k,v in todo.items() if k not in out}
print('todo',len(todo),flush=True)
lock=threading.Lock(); n=[0]
def work(k):
    h=subprocess.run(['curl','-sSL','-m','40','-A','Mozilla/5.0','https://www.saudijobs24.com/'+k+'-topic'],capture_output=True).stdout.decode('utf-8','ignore')
    t=re.sub(r'<script.*?</script>|<style.*?</style>','',h,flags=re.S); t=H.unescape(re.sub(r'\s+',' ',re.sub(r'<[^>]+>',' ',t)))
    i=t.find('بواسطة: وظائف السعودية 24')
    body=t[i+25:i+2500] if i>=0 else ''
    j=t.rfind('وظائف شاغرة في السعودية',0,i) if i>=0 else -1
    title=t[j+24:i].strip()[:250] if j>=0 else ''
    d=re.search(r'(\d{2}) - (\d{2}) - (\d{4})',t[max(0,i-40):i+5]) if i>=0 else None
    with lock:
        out[k]=dict(tg=tg_id[k],title=title,date=f'{d.group(3)}-{d.group(2)}-{d.group(1)}' if d else '',body=body)
        n[0]+=1
        if n[0]%250==0:
            json.dump(out,open('sj_items.json','w'),ensure_ascii=False); print('sj',n[0],flush=True)
    time.sleep(0.1)
tg_id=todo
with ThreadPoolExecutor(8) as ex: list(ex.map(work,list(todo)))
json.dump(out,open('sj_items.json','w'),ensure_ascii=False)
print('done sj',len(out),flush=True)
