import json, re, collections, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

T1 = re.compile(r'تحليل البيانات|تحليل بيانات|محلل بيانات|محلل أعمال|محلل نظم|محلل|إحصائي|احصائي|إحصاء|احصاء|ذكاء الأعمال|ذكاء اعمال|Power ?BI|\bSQL\b|Tableau|\bdata\b|analyst|analytics|\bBI\b|statistic|informatics|معلوماتية|علم البيانات|إدارة البيانات|ادارة البيانات|هندسة البيانات|مهندس بيانات|أخصائي بيانات|اخصائي بيانات|جودة البيانات|حوكمة البيانات|الأرشفة|أرشفة|\bMIS\b|reporting|أخصائي تقارير|اخصائي تقارير', re.I)
T2 = re.compile(r'تقنية المعلومات|تقنية معلومات|نظم المعلومات|نظم معلومات|\bIT\b|برمج|مطور|developer|ERP|SAP|Oracle|قواعد البيانات|database|أمن سيبراني|الأمن السيبراني|cyber|حاسب آلي|حاسب الي|علوم الحاسب|computer|شبكات|\bnetwork|\bsystems? (?:analyst|admin|engineer)|دعم فني تقني|help ?desk', re.I)
EXCL = re.compile(r'إدخال بيانات|ادخال بيانات|مدخل بيانات|مدخلة بيانات|data entry', re.I)

def tier(text):
    if T1.search(EXCL.sub('', text)):
        return 1
    if T2.search(text):
        return 2
    return 3

rows = []  # source, company, title, date, url, text, tier

# tanqeeb
if os.path.exists('tanq_items.json'):
    for v in json.load(open('tanq_items.json')).values():
        txt = ' '.join([v.get('name', ''), v.get('name-alt', ''), v.get('categories', ''), v.get('industry', '')])
        rows.append(dict(src='Tanqeeb (' + v.get('source', '') + ')', company=v.get('company', ''), title=v.get('name', ''),
                         date=v.get('date', ''), url='https://saudi.tanqeeb.com' + v.get('url', ''), text=txt, tier=tier(txt),
                         industry=v.get('industry', ''), exp=v.get('experience', '')))

# sabbar companies
if os.path.exists('comp_taif.json'):
    for sl, v in json.load(open('comp_taif.json')).items():
        for u, n in v['taif_jobs']:
            rows.append(dict(src='Sabbar (صفحة الشركة)', company=v['name'], title=n, date='', url=u, text=n + ' ' + u,
                             tier=tier(n + ' ' + re.sub(r'.*/c-taif-r-|/id-.*', '', u).replace('-', ' ')), industry='', exp='',
                             comp_url='https://sabbar.com/en/jobs/companies/' + sl, desc=v.get('desc', '')))

# wdeftksa (Taif-specific only)
if os.path.exists('wdeft_items.json'):
    for v in json.load(open('wdeft_items.json')):
        loc = v.get('loc', '')
        if 'جميع' in loc or not loc:
            continue
        if 'الطائف' not in loc and 'الطائف' not in v['title']:
            continue
        if len(loc.split('،')) > 3:
            continue
        txt = v['title'] + ' ' + ' '.join(v.get('tags', []))
        rows.append(dict(src='وظيفتك علينا', company=v['company'], title=v['title'], date=v['date'], url=v['url'], text=txt,
                         tier=tier(txt), industry='', exp=''))

# telegram
for f in os.listdir('.'):
    if f.startswith('tg_') and f.endswith('.json'):
        for v in json.load(open(f)).values():
            t = v['text']
            if 'الطائف' not in t or '#عن_بعد' in t:
                continue
            first = t.split('\n')[0][:200]
            m = re.search(r'(?:لدى|في|بـ|ب)\s*((?:شركة|مجموعة|مؤسسة|مستشفى|مجمع|مصنع|مدارس|جامعة|هيئة|أمانة|جمعية|بنك|مصرف|مركز|معهد|صيدليات|أسواق|فندق|مطاعم)[^\n،,\(\)|▪️🔗📍]{2,60})', t)
            comp = m.group(1).strip() if m else ''
            rows.append(dict(src='تليجرام وظائف الطائف 24', company=comp, title=first, date=v['date'][:10],
                             url='https://t.me/jobstaif/%s' % v['id'], text=t[:800], tier=tier(t[:800]), industry='', exp=''))

json.dump(rows, open('all_rows.json', 'w'), ensure_ascii=False)
c = collections.Counter((r['tier']) for r in rows)
print('rows', len(rows), c)
comp = collections.defaultdict(lambda: collections.Counter())
for r in rows:
    comp[r['company'].strip()][r['tier']] += 1
lst = sorted(comp.items(), key=lambda kv: (-kv[1][1], -kv[1][2], -sum(kv[1].values())))
print('companies', len(lst))
for k, v in lst[:150]:
    if v[1] or v[2]:
        print(dict(v), k)
