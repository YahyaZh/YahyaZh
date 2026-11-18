import json, re, glob, os, collections
os.chdir(os.path.dirname(os.path.abspath(__file__)))
exec(open('analyze.py', encoding='utf-8').read().split("rows = []")[0])  # tier()

VERB = r'\s+(?:توفر|تُوفر|يوفر|تعلن|تُعلن|يعلن|أعلنت|اعلنت|تطلق|تفتح|تبدأ|بدأت|يطلق|تطرح|تقدم|يقدم|توظف|لديها|بحاجة|عن|براتب|برواتب|بمكافأة|للعمل|بالطائف|في الطائف|بمدينة|-|–|:|\||\(|،|,)'
ORGP = r'((?:الشركة|شركة|مجموعة|مؤسسة|مستشفى|مستشفيات|مجمع|مصنع|مدارس|هيئة|الهيئة|جمعية|مركز|معهد|صيدليات|أسواق|فندق|مطاعم|جامعة|أمانة|وزارة|بنك|مصرف|البنك)\s[^\n،,\(\)▪️🔗📍|:]{2,50})'

def clean(n):
    n = re.sub(r'\s+', ' ', n or '').strip()
    n = re.sub(r'^(جامعة|أمانة|امانة|أمانة محافظة|جمعية شباب|هيئة تطوير محافظة|جمعية الدعوة بمحافظة)$', lambda m: m.group(0)+' الطائف', n)
    n = re.split(VERB, ' ' + n + ' ')[0].strip() if re.search(VERB, ' ' + n + ' ') else n
    n = re.sub(r'\s*[-–]\s*الطائف.*$', '', n)
    n = re.sub(r'(بالطائف|في الطائف|الطائف)$', lambda m: m.group(0) if n.startswith(('شركة الطائف', 'مستشفى الطائف', 'مجمع الطائف', 'مصنع الطائف', 'جمعية الطائف')) else '', n).strip()
    n = re.sub(r'\s*(رواتب|براتب|برواتب).*$', '', n).strip(' .:-')
    n = re.sub(r'\s+(بمحافظة|بمدينة|عبر)$', '', n)
    n = re.sub(r'^(جامعة|أمانة|امانة|أمانة محافظة|هيئة تطوير محافظة|جمعية شباب|جمعية الدعوة بمحافظة|جمعية الدعوة)$', lambda m: m.group(0)+' الطائف', n)
    if re.match(r'^جامعة الطائف', n): n='جامعة الطائف'
    if re.match(r'^(أمانة|امانة) (محافظة )?الطائف', n): n='أمانة محافظة الطائف'
    if re.match(r'^شركة شبرا', n): n='شركة شبرا الطائف التجارية'
    if re.match(r'^شركة الافق', n): n='شركة الأفق لتقنية المعلومات'
    if re.match(r'^شركة العرض المتقن', n): n='شركة العرض المتقن للخدمات التجارية'
    if re.match(r'^(مستشفيات القوات المسلحة|الخدمات الطبية للقوات المسلحة)', n): n='مستشفيات القوات المسلحة بالطائف'
    if re.match(r'^شركة (سراكو|ساراكو)', n): n='شركة سراكو'
    if re.match(r'^(شركة )?وقف جامعة', n) or re.match(r'^شركة الوقف بجامعة', n): n='شركة وقف جامعة الطائف'
    return n

def key(n):
    k = re.sub(r'^(ال)?(شركة|مؤسسة|مجموعة)\s+', '', n)
    k = re.sub(r'\s+(المحدودة|المحدوده|شركة شخص واحد|شخص واحد|التجارية|التجاريه|القابضة)\b', '', k)
    k = re.sub(r'[ًٌٍَُِّْـ]', '', k)
    k = k.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ة', 'ه').replace('ى', 'ي')
    return re.sub(r'[^\w]', '', k).lower()

recs = []
NONJOB = re.compile(r'القبول|المقبولين|دورة|دورات|ورشة|ورش |محاضرة|فعالية|الدراسات العليا|برامج الماجستير|للماجستير|ابتعاث|نتائج|اختبار|التطوع|تطوعية|منحة|مسابقة|ملتقى|معرض|استبيان')
def add(src, org, title, date, url, text, scope):
    if NONJOB.search(title or ''):
        return
    org = clean(org)
    if not org or len(org) < 4 or org in ('شركة', 'مؤسسة', 'مجمع طبي', 'شركة كبرى', 'شركة عقارية', 'مجموعة'):
        return
    if re.match(r'^(شركة|مؤسسة|مجمع|مركز|مصنع|مدارس|فندق)\s+(كبرى|رائدة|خاصة|خاص|طبي|متخصص|وطنية|في|للمقاولات|مشغلة|مالية|عقارية|ملابس|حلويات|استشارات|راقي|للاطفال|أطفال|بقسم|للطب|عالمية|كبرى)', org):
        return
    recs.append(dict(src=src, org=org, k=key(org), title=title.strip()[:200], date=date, url=url, tier=tier(text[:900]), scope=scope))

# Telegram
tg = {}
for f in glob.glob('tg_jobstaif*.json'):
    tg.update(json.load(open(f)))
for v in tg.values():
    t = v['text']
    if 'الطائف' not in t or '#عن_بعد' in t or re.search(r'دورة|دورات|ورشة|ندوة|ندوات|استبيان|بحث يهدف|القبول|whatsapp|واتساب|pinned', t[:300]):
        continue
    m = re.search(r'(?:في|لدى|بـ|ب|عنها|:)\s*' + ORGP, t) or re.search(ORGP, t)
    if not m:
        continue
    scope = 'الطائف' if ('#الطائف' in t or re.search(r'الطائف\s*$', t.split('\n')[0]) or t.count('،') < 4) else 'عدة مدن منها الطائف'
    add('تليجرام: وظائف الطائف 24', m.group(1), t.split('\n')[0], v['date'][:10], 'https://t.me/jobstaif/%s' % v['id'], t, scope)

# Tanqeeb (current, Taif-located)
for v in json.load(open('tanq_items.json')).values():
    txt = ' '.join([v.get('name', ''), v.get('name-alt', ''), v.get('categories', '')])
    add('Tanqeeb / ' + v.get('source', ''), v.get('company', ''), v.get('name', ''), '2026-11 (حالي)', v.get('share-url', ''), txt, 'الطائف')

# Sabbar company pages
for sl, v in json.load(open('comp_taif.json')).items():
    for u, n in v['taif_jobs']:
        add('Sabbar', v['name'], n, '2026 (حالي/حديث)', u, n + ' ' + u.replace('-', ' '), 'الطائف')

# wdeftksa
for v in json.load(open('wdeft_items.json')):
    loc = v.get('loc', '')
    if 'جميع' in loc or 'الطائف' not in (loc + v['title']):
        continue
    nc = len(loc.split('،'))
    if nc > 6:
        continue
    add('وظيفتك علينا', v['company'] or v['title'], v['title'], v['date'], v['url'], v['title'] + ' ' + ' '.join(v.get('tags', [])),
        'الطائف' if nc <= 2 else 'عدة مدن منها الطائف')

# ewdifh
if os.path.exists('ew_items.json'):
    for v in json.load(open('ew_items.json')).values():
        ti, tx = v['title'], v['text']
        if 'الطائف' not in ti and not re.search(r'\(الطائف\)|بالطائف|في الطائف|الطائف\)', tx):
            continue
        if re.search(r'جميع مناطق|كافة مناطق', ti):
            continue
        m = re.search(ORGP, ti)
        org = m.group(1) if m else ''
        add('أي وظيفة (ewdifh)', org, ti, '-'.join(reversed(v['date'].split('-'))) if v['date'] else '', 'https://www.ewdifh.com/jobs/' + v['id'], ti + ' ' + tx[:1500],
            'الطائف' if 'الطائف' in ti and ti.count('،') < 3 else 'عدة مدن منها الطائف')

# aggregate
agg = collections.OrderedDict()
for r in recs:
    a = agg.setdefault(r['k'], dict(names=collections.Counter(), n=0, t1=0, t2=0, dates=[], srcs=set(), ex1=[], ex2=[], ex=[], taifonly=0))
    a['names'][r['org']] += 1
    a['n'] += 1
    a['srcs'].add(r['src'].split(' / ')[0])
    if r['date']:
        a['dates'].append(r['date'])
    if r['scope'] == 'الطائف':
        a['taifonly'] += 1
    item = (r['date'], r['title'], r['url'])
    if r['tier'] == 1:
        a['t1'] += 1; a['ex1'].append(item)
    elif r['tier'] == 2:
        a['t2'] += 1; a['ex2'].append(item)
    else:
        a['ex'].append(item)

out = []
for k, a in agg.items():
    name = a['names'].most_common(1)[0][0]
    out.append(dict(key=k, name=name, n=a['n'], t1=a['t1'], t2=a['t2'], taifonly=a['taifonly'], srcs=sorted(a['srcs']),
                    first=min(a['dates']) if a['dates'] else '', last=max(a['dates']) if a['dates'] else '',
                    ex1=a['ex1'][:4], ex2=a['ex2'][:4], ex=a['ex'][:4]))
json.dump(out, open('employers.json', 'w'), ensure_ascii=False, indent=0)
json.dump(recs, open('recs.json', 'w'), ensure_ascii=False)
print('records', len(recs), 'employers', len(out))
print('with t1', sum(1 for o in out if o['t1']), 'with t2', sum(1 for o in out if o['t2']))
