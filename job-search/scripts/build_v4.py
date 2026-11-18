import sys, json, re, os
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1]
src = open(os.path.join(HERE, 'build_xlsx.py'), encoding='utf-8').read()
data_part = src.split('# ---------------------------------------------------------------- styles')[0]
data_part = data_part.replace('OUT = sys.argv[1]', '')
exec(data_part)
exec(open(os.path.join(HERE, 'curated_v2.py'), encoding='utf-8').read())
exec(open(os.path.join(HERE, 'curated_v3.py'), encoding='utf-8').read())

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

sys.path.insert(0, HERE)
from domains import domains_of_ad, DOMAINS, DE_RX, IT_RX
DOM_NAMES = [d[0] for d in DOMAINS]
recs = json.load(open(os.path.join(HERE, 'recs.json'), encoding='utf-8'))
emps = json.load(open(os.path.join(HERE, 'employers.json'), encoding='utf-8'))

# manual job records found through targeted web searches (not in scraped archives)
MANUAL_JOBS = [
    ("2025–2026", "BAE Systems Saudi Arabia", "Data Management Technician (Level 1)", "https://jobsearch.baesystems.com/job/data-management-technician-level-1-00132782", 1, "الطائف", "بحث Google"),
    ("2025–2026", "BAE Systems Saudi Arabia", "In-Service Data Analysis (ISDA) Specialist", "https://jobsearch.baesystems.com/job/in-service-data-analysis-isda-specialist-00134354", 1, "الطائف", "بحث Google"),
    ("2025–2026", "BAE Systems Saudi Arabia", "Senior Technical Specialist – Data Integrity", "https://sabbar.com/en/jobs/c-taif-r-senior-technical-specialist-data-integrity/id-fbd0c7c2-6717-41fd-a274-19d4b65ac1ff", 1, "الطائف", "بحث Google"),
    ("2025–2026", "BAE Systems Saudi Arabia", "ESS Business Data Senior Specialist", "https://sabbar.com/en/jobs/c-taif-r-ess-business-data-senior-specialist/id-8d50462b-7f27-4c85-a453-fd01c60bfeda", 1, "الطائف", "بحث Google"),
    ("2025-02-03", "BAE Systems Saudi Arabia", "Senior Data Analyst Specialist", "https://jobs.y-axis.com/jobs/senior-data-analyst-specialist-bae-systems", 1, "الطائف", "بحث Google"),
    ("2026-06", "هيئة تطوير محافظة الطائف", "مدير إدارة البيانات (جدارات)", "https://www.saudijobs24.com/t180709-topic", 1, "الطائف", "بحث Google"),
    ("2026-06-15", "شركة غير مسمّاة تعمل مع المياه الوطنية", "موظفة تحليل بيانات", "https://www.jobs-arab.com/sa/jobs/1149772002/", 1, "الطائف", "بحث Google"),
    ("2025-07", "جمعية الدعوة والإرشاد بالطائف", "محلل بيانات وتقارير المدير التنفيذي (بطاقة وصف وظيفي)", "https://dawahtaif.org.sa/wp-content/uploads/sites/596/2025/07/%D8%AF%D9%84%D9%8A%D9%84-%D8%A7%D9%84%D8%A3%D9%88%D8%B5%D8%A7%D9%81-%D8%A7%D9%84%D9%88%D8%B8%D9%8A%D9%81%D9%8A%D8%A9.pdf", 1, "الطائف", "بحث Google"),
    ("غير محدد", "شركة الأفق لتقنية المعلومات", "أخصائي جودة (فهرسة وتدقيق الأرشيف) – 4000", "https://sabbar.com/en/jobs/c-taif-r-quality-specialist/id-16e10025-1fd0-4cb0-90a5-becec236bcdd", 1, "الطائف", "بحث Google"),
    ("غير محدد", "شركة الأفق لتقنية المعلومات", "مدير إدارة مشاريع – 8000", "https://sabbar.com/ar/jobs/c-الطائف-r-مدير-إدارة-مشاريع/id-385c6146-a6af-4bc6-8a99-6026417abcc4", 3, "الطائف", "بحث Google"),
    ("2026", "مجموعة المجال العربي", "فني شبكات تقنية معلومات – 7000", "https://sabbar.com/ar/jobs/c-الطائف-r-فني-شبكات-تقنية-معلومات/id-876b7c56-e914-456a-bd03-39c3b74df575", 2, "الطائف", "بحث Google"),
    ("2026", "شركة سراكو", "IT Support Technician – 6000 (0–1 سنة)", "https://sabbar.com/en/jobs/m-technical-support-f-for-fresh-graduates", 2, "الطائف", "بحث Google"),
    ("غير محدد", "مجموعة بن شيهون", "Quality Technician", "https://sabbar.com/en/jobs/c-taif-r-quality-technician/id-c3a2b3ff-42c7-4455-869d-56f850936c0f", 3, "الطائف", "بحث Google"),
    ("غير محدد", "صيدليات الدواء", "Supply Chain Specialist", "https://sabbar.com/en/jobs/c-taif-r-supply-chain-specialist/id-228b8c1e-3ae4-4fc6-bd91-70d4b5ceabf7", 3, "الطائف", "بحث Google"),
    ("2021-08-31", "شركة العرض المتقن للخدمات التجارية", "مطوّر تطبيقات الجوال (الطائف) – عبر LinkedIn", "https://www.ewdifh.com/jobs/68116", 2, "الطائف", "بحث Google"),
    ("2021", "شركة العرض المتقن للخدمات التجارية", "أخصائي شبكات / مهندس نظم / مهندس حاسب (الطائف)", "https://www.ewdifh.com/jobs/57861", 2, "عدة مدن منها الطائف", "بحث Google"),
    ("غير محدد", "شركة توت المصيف للتجارة", "أخصائي تسويق / أمين صندوق (تحليل التدفقات النقدية)", "https://sabbar.com/en/jobs/c-taif-r-treasurer/id-8b94d58f-ca0c-42ce-b6f9-97878e91bff7", 3, "الطائف", "بحث Google"),
    ("غير محدد", "جمعية البر الخيرية بالطائف", "أخصائي تقنية معلومات (مغلق)", "https://www.wdeftksa.com/sa/jobs/100080", 2, "الطائف", "بحث Google"),
    ("غير محدد", "جمعية الزواج والتنمية الأسرية بالطائف", "أخصائي تقنية معلومات", "https://www.wdeftksa.com/sa/jobs/3472", 2, "الطائف", "بحث Google"),
    ("غير محدد", "مستشفى النهضة العام", "مهندس تقنية معلومات", "https://www.zavis.ai/ar/sa/directory/taif/hospitals/al-nahda-general-hospital", 2, "الطائف", "بحث Google"),
    ("2025–2026", "BAE Systems Saudi Arabia", "Engineering Support System (ESS) Reference Data Senior Specialist (بيانات مرجعية – فريق سلامة البيانات)", "https://sabbar.com/en/jobs/c-taif-r-engineering-support-system-ess-reference-data-senior-specialist/id-2452c181-4f87-40ee-b700-557f09c3abce", 1, "الطائف", "بحث Google"),
    ("2025–2026", "BAE Systems Saudi Arabia", "Engineering Support System (ESS) National Data Specialist (للكوادر الوطنية)", "https://jobsearch.baesystems.com/job/ess-national-data-specialist-00134513", 1, "الطائف", "بحث Google"),
    ("2025–2026 (غير مؤكد)", "Schneider Electric", "Digital Transformation Data Engineer with AI & Automation (مصدر واحد يذكر الطائف وآخر يذكر الرياض)", "https://www.jobleads.com/sa/job/digital-transformation-data-engineer-with-ai-automation--at-ta-if--e977aee51eccfe2dee88452492aeae5a5", 1, "غير مؤكد", "بحث Google"),
    ("2025-04-16", "غير مسمّى (Jobmel)", "مهندس معالجة بيانات – الطائف (تحليل بيانات وقواعد بيانات وبرمجة)", "https://ar.jobmel.com/Saudi-Arabia/Taif/Data-processing-engineer-30", 1, "الطائف", "بحث Google"),
    ("2026-11", "الغرفة التجارية الصناعية بالطائف", "أخصائي أمن سيبراني – 7000", "https://sabbar.com/ar/jobs/c-الطائف-r-أخصائي-أمن-سيبراني/id-81897e39-9f77-4196-9624-07ecadd3a8d0", 2, "الطائف", "Sabbar"),
]

DE = re.compile(r'إدخال بيانات|إدخال البيانات|ادخال بيانات|ادخال البيانات|مدخل بيانات|مدخلي بيانات|مدخلين بيانات|مدخل/ة بيانات|مدخلة بيانات|مدخلات بيانات|data entry', re.I)
def cat(tier, title, doms=None):
    if DE.search(title):
        return "إدخال بيانات (مستبعد)"
    if doms:
        return "بيانات / تحليل / إحصاء"
    if doms is not None and tier == 1:
        tier = 3
    if DE.search(title):
        return "إدخال بيانات (مستبعد)"
    return {1: "بيانات / تحليل / إحصاء", 2: "تقنية معلومات", 3: "أخرى"}[tier]

def status(d):
    d = str(d)
    if 'حالي' in d or d >= '2026-09':
        return "حالي/حديث (تحقق)"
    return "سابق"

def sector(name):
    rules = [(r'BAE|سيستمز|Systems|طيران|الدفاع|الحرس|الداخلية', 'دفاع وأمن وطيران'), (r'مياه|كهرباء|الطاقة|الري', 'مرافق وطاقة'), (r'الأهلية|الاهلية|العبادي', 'تعليم'), (r'مستشفى|مستشفيات|طبي|طبية|صحي|صحية|صيدلي|أسنان|تجمع|جراح|علاج|عيادات|عيادة|رعاية|حجامة|تغذية|نبض الحياة', 'صحة'), (r'جامعة|مدارس|مدرسة|تعليم|أكاديمية|معهد|كلية', 'تعليم'),
             (r'جمعية|وقف|خيري', 'غير ربحي'), (r'الجيل النموذجية|تنمية الإنسان|المراس', 'تعليم'), (r'أمانة|هيئة|وزارة|الإحصاء|ديوان|المؤسسة العامة|الغرفة', 'حكومي/شبه حكومي'),
             (r'تقنية|معلومات|برمج|سوفت|soft|tech|ماتركس|الأفق', 'تقنية'), (r'مقاولات|إنشاء|انشاء|هندس|معمار', 'مقاولات/هندسة'),
             (r'مصنع|صناع|ألبان|البان|أغذية|اغذية|غذائية|مخابز', 'صناعة/أغذية'), (r'مطاعم|مطعم|قهوة|كافيه|كوفي|برجر|البيك', 'مطاعم ومقاهي'),
             (r'فندق|فنادق|سياح|ضيافة|إنتركونتيننتال', 'ضيافة وسياحة'), (r'تجار|تجزئة|أسواق|اسواق|بنده|سوبر|أباريل|عطور|مفروشات|هوم|المراعي|جميل', 'تجزئة وتجارة'),
             (r'بنك|مصرف|مالية|تمويل', 'مالية'), (r'نقل|لوجست|شحن|توصيل', 'نقل ولوجستيات'), (r'تشغيل|صيانة|مرافق|سراكو', 'تشغيل وصيانة')]
    for p, s in rules:
        if re.search(p, name, re.I):
            return s
    return 'أخرى'

# ------------------------------------------------ styles
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
BODY_FONT = Font(name="Arial", size=10)
LINK_FONT = Font(name="Arial", size=10, color="0563C1", underline="single")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="right", readingOrder=2)
CENTER = Alignment(wrap_text=True, vertical="top", horizontal="center", readingOrder=2)
STATUS_LIST = '"لم أقدّم,قدّمت إلكترونياً,قدّمت يدوياً,تواصلت مع شخص,مقابلة,عرض,مرفوض,مؤجل"'

wb = Workbook()

def sheet(title, headers, widths, rows, link_cols=(), center_cols=(), table_name=None, row_h=None, status_col=None, freeze="B2"):
    s = wb.create_sheet(title)
    s.sheet_view.rightToLeft = True
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = s.cell(row=1, column=ci, value=h)
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CENTER, BORDER
        s.column_dimensions[get_column_letter(ci)].width = w
    s.row_dimensions[1].height = 34
    for ri, r in enumerate(rows, 2):
        for ci, v in enumerate(r, 1):
            c = s.cell(row=ri, column=ci, value=v)
            c.font, c.border = BODY_FONT, BORDER
            c.alignment = CENTER if ci in center_cols else WRAP
            if ci in link_cols and isinstance(v, str) and v.startswith("http"):
                c.hyperlink = v
                c.font = LINK_FONT
        if row_h:
            s.row_dimensions[ri].height = row_h
    last = max(len(rows) + 1, 2)
    if table_name:
        t = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(headers))}{last}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        s.add_table(t)
    else:
        s.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last}"
    if status_col:
        L = get_column_letter(status_col)
        dv = DataValidation(type="list", formula1=STATUS_LIST, allow_blank=True)
        s.add_data_validation(dv)
        dv.add(f"{L}2:{L}3000")
    s.freeze_panes = freeze
    return s

# ------------------------------------------------ sheet: curated entities

GOV_DATA_CONTRACTORS = re.compile(r'BAE|سراكو|SRACO|تحكم|Tahakom|بكتل|Bechtel|معادن|المجال العربي|العرض المتقن|السلام لصناعة|Alsalam', re.I)
def ndmo_of(r):
    t = f"{r.get('typ','')} {r.get('ar','')} {r.get('en','')}"
    if re.search(r'شبه خاص', t):
        return "نظام حماية البيانات الشخصية: إلزامي. ضوابط NDMO: إرشادية، وتصير إلزامية على أي بيانات حكومية تديرها الجهة (تحقق من النص)"
    if re.search(r'حكومي|شبه حكومي|غرفة', t):
        return "إلزامي: ضوابط NDMO (مكتب لإدارة البيانات، ومسؤول بيانات، وقياس امتثال سنوي) + نظام حماية البيانات الشخصية"
    if GOV_DATA_CONTRACTORS.search(t):
        return "نظام حماية البيانات الشخصية: إلزامي. ضوابط NDMO: تنطبق على البيانات الحكومية اللي تديرها الجهة كمقاول (حسب ملخصات الضوابط، تحقق من النص الرسمي)"
    if re.search(r'غير ربحي|جمعية', t):
        return "نظام حماية البيانات الشخصية: إلزامي. ضوابط NDMO: إرشادية (مرجع جيد لحوكمة الجمعية)"
    return "نظام حماية البيانات الشخصية: إلزامي على كل من يعالج بيانات شخصية. ضوابط NDMO: إرشادية للقطاع الخاص"
def pot_domains(r):
    t = f"{r.get('sector','')} {r.get('typ','')} {r.get('ar','')} {r.get('roles','')} {r.get('dept','')}"
    out = []
    if re.search(r'حكومي|غرفة|هيئة|أمانة|جامعة', t): out += ["حوكمة/جودة/إدارة (DMO)", "تحليل وBI"]
    if re.search(r'صحة|طبي|مستشفى|تجمع', t): out += ["إحصاء/معلوماتية صحية", "تحليل وBI", "حوكمة/جودة/إدارة (DMO)"]
    if re.search(r'دفاع|BAE|طيران', t): out += ["حوكمة/جودة/إدارة (DMO)", "تحليل وBI"]
    if re.search(r'تقنية معلومات|تقنية المعلومات|برمجة|\bIT\b|Software|\bTech\b', t): out += ["هندسة بيانات/ETL/قواعد", "علم بيانات/AI"]
    if re.search(r'أرشفة|وثائق|قانون|محاماة|Integrity|بيانات مرجعية', t): out += ["وثائق وسجلات", "حوكمة/جودة/إدارة (DMO)"]
    if re.search(r'تجزئة|تجارة|صناعة|غذائية|ضيافة|فنادق|مطاعم|تطوير عقاري', t): out += ["تحليل وBI", "هندسة بيانات/ETL/قواعد"]
    if re.search(r'تعدين|مشاريع|مقاولات|تشغيل', t): out += ["تحليل وBI", "وثائق وسجلات"]
    if re.search(r'إحصاء', t): out += ["إحصاء/معلوماتية صحية"]
    seen=[]; [seen.append(x) for x in out if x not in seen]
    return "، ".join(seen) or "تحليل وBI"
for r in R:
    r['ndmo'] = ndmo_of(r)
    r['pdom'] = pot_domains(r)
COLS.insert(6, ("مكتب إدارة البيانات (NDMO) وحماية البيانات", 34))
COLS.insert(7, ("مجالات البيانات المحتملة فيها", 30))
keys = ["_id", "p", "ar", "en", "sector", "typ", "ndmo", "pdom", "size", "loc", "has", "ev", "evd", "date", "roles", "dept",
        "fit", "cond", "how", "jobs", "web", "li", "x", "mail", "contact", "found", "s1", "s2", "s3", "nxt",
        "_status", "_date", "_notes"]
order = {"أ": 0, "ب": 1, "ج": 2}
R.sort(key=lambda r: (order.get(r['p'], 3), 0 if str(r['has']).startswith('مؤكد') else 1))
cur_rows = []
for i, r in enumerate(R, 1):
    r['_id'] = i
    r['_status'] = 'لم أقدّم'
    cur_rows.append([r.get(k, '') for k in keys])
hdrs = [c[0] for c in COLS]
wds = [c[1] for c in COLS]
li = {keys.index(k) + 1 for k in ("jobs", "web", "li", "x", "s1", "s2", "s3")}
s1 = sheet("الجهات ذات الأولوية", hdrs, wds, cur_rows, link_cols=li, center_cols={1, 2, 11, 17, 31, 32},
           table_name="Priority", row_h=150, status_col=keys.index("_status") + 1, freeze="D2")
has_col = get_column_letter(keys.index("has") + 1); pri_col = get_column_letter(2)
for txt, color in [("مؤكد", "C6EFCE"), ("محتمل قوي", "DDEBF7"), ("محتمل", "FFF2CC"), ("ممكن", "F2F2F2")]:
    s1.conditional_formatting.add(f"{has_col}2:{has_col}500", FormulaRule(formula=[f'LEFT({has_col}2,{len(txt)})="{txt}"'], fill=PatternFill("solid", fgColor=color), stopIfTrue=True))
for txt, color in [("أ", "63BE7B"), ("ب", "FFEB84"), ("ج", "D9D9D9")]:
    s1.conditional_formatting.add(f"{pri_col}2:{pri_col}500", FormulaRule(formula=[f'{pri_col}2="{txt}"'], fill=PatternFill("solid", fgColor=color)))

# ------------------------------------------------ sheet: jobs (current + past)
job_rows = []
seen = set()
for (d, org, title, url, tier, scope, srcn) in MANUAL_JOBS:
    _dm = domains_of_ad(title) or (['تحليل البيانات وذكاء الأعمال'] if tier == 1 and re.search(r'Analys|تحليل|محلل', title, re.I) else (['حوكمة وإدارة وجودة البيانات (DMO)'] if tier == 1 else []))
    job_rows.append([d, status(d), org, sector(org), title, cat(tier, title, _dm), '، '.join(_dm), scope, srcn, url] + ['✓' if n in _dm else '' for n in DOM_NAMES])
    seen.add(url)
for r in recs:
    if r['url'] in seen:
        continue
    seen.add(r['url'])
    _dm = r.get('doms', [])
    job_rows.append([r['date'] or 'غير محدد', status(r['date']), r['org'], sector(r['org']), r['title'], cat(r['tier'], r['title'], _dm), '، '.join(_dm), r['scope'], r['src'], r['url']] + ['✓' if n in _dm else '' for n in DOM_NAMES])
catorder = {"بيانات / تحليل / إحصاء": 0, "تقنية معلومات": 1, "أخرى": 2, "إدخال بيانات (مستبعد)": 3}
job_rows.sort(key=lambda x: (catorder[x[5]], '' if x[0] == 'غير محدد' else x[0]), reverse=False)
job_rows.sort(key=lambda x: catorder[x[5]])
for i, r in enumerate(job_rows, 1):
    r.insert(0, i)
    r += ['', '']
SHORT = ["تحليل وBI", "هندسة بيانات/ETL/قواعد", "علم بيانات/AI", "حوكمة/جودة/إدارة (DMO)", "إحصاء/معلوماتية صحية", "وثائق وسجلات"]
jh = ["م", "تاريخ الإعلان", "حالي / سابق", "الجهة", "القطاع (تقديري)", "عنوان الإعلان / المسمى", "التصنيف", "مجالات البيانات", "نطاق المكان", "المصدر", "رابط الإعلان"] + SHORT + ["هل يهمني؟", "ملاحظاتي"]
jw = [6, 14, 16, 34, 16, 70, 20, 30, 18, 22, 46] + [12]*6 + [12, 26]
s2 = sheet("الوظائف (حالية وسابقة)", jh, jw, job_rows, link_cols={11}, center_cols={1, 2, 3, 7, 9, 12, 13, 14, 15, 16, 17, 18}, table_name="Jobs", row_h=None, freeze="E2")
dv = DataValidation(type="list", formula1='"نعم,ربما,لا"', allow_blank=True); s2.add_data_validation(dv); dv.add("R2:R5000")
for txt, color in [("بيانات", "C6EFCE"), ("تقنية", "DDEBF7"), ("إدخال", "F8CBAD")]:
    s2.conditional_formatting.add("G2:G5000", FormulaRule(formula=[f'LEFT(G2,{len(txt)})="{txt}"'], fill=PatternFill("solid", fgColor=color)))

# ------------------------------------------------ sheet: all employers
curated_names = [r['ar'] for r in R]
def in_cur(n):
    k = re.sub(r'^(شركة|مؤسسة|مجموعة)\s+', '', n)[:10]
    return any(k and k in c for c in curated_names)
emp_rows = []
for o in sorted(emps, key=lambda o: (-o['t1'], -o['t2'], -o['n'])):
    ex = (o['ex1'] + o['ex2'] + o['ex'])[:3]
    titles = "\n".join(f"• {e[1][:90]} ({e[0] or '؟'})" for e in ex)
    links = [e[2] for e in ex] + ['', '', '']
    _dsum = '، '.join(f"{k} ({v})" for k, v in sorted(o.get('dom', {}).items(), key=lambda kv: -kv[1]))
    emp_rows.append([o['name'], sector(o['name']), o['n'], o['t1'], o['t2'], _dsum, o['taifonly'], o['first'] or '', o['last'] or '',
                     "، ".join(o['srcs']), titles, links[0], links[1], links[2], 'نعم' if in_cur(o['name']) else '', 'لم أقدّم', ''])
for i, r in enumerate(emp_rows, 1):
    r.insert(0, i)
eh = ["م", "الجهة (كما وردت في الإعلانات)", "القطاع (تقديري)", "عدد الإعلانات في الطائف", "إعلانات بيانات", "إعلانات تقنية معلومات", "مجالات البيانات في إعلاناتها",
      "إعلانات خاصة بالطائف فقط", "أقدم إعلان", "أحدث إعلان", "المصادر", "أمثلة من إعلاناتها", "رابط 1", "رابط 2", "رابط 3", "موجودة في ورقة الأولوية؟", "حالة التقديم", "ملاحظاتي"]
ew = [6, 38, 16, 12, 12, 12, 30, 12, 13, 13, 26, 70, 30, 30, 30, 12, 14, 26]
s3 = sheet("كل الجهات المعلنة في الطائف", eh, ew, emp_rows, link_cols={13, 14, 15}, center_cols={1, 4, 5, 6, 8, 9, 10, 16, 17},
           table_name="AllEmployers", row_h=60, status_col=17, freeze="C2")


exec(open(os.path.join(HERE, 'manual_sheet.py'), encoding='utf-8').read())

# ------------------------------------------------ employee evidence
EVID = [
    ("تجمع الطائف الصحي", "Associate Executive Director of Business Intelligence (منذ 2024)، وكانت قبلها Data Specialist", "إدارة BI فعلية", "https://www.linkedin.com/in/hanan-a-b70296160/"),
    ("تجمع الطائف الصحي", "خريجة نظم معلومات إدارية من جامعة الطائف تشتغل بـ Power BI وSQL", "تحليل بيانات داخل التجمع", "https://www.linkedin.com/in/wafa-ajlan-alkhaildi-7a0038401/"),
    ("تجمع الطائف الصحي", "Quality Officer (منذ 2024)", "الجودة ومؤشرات الأداء", "https://www.linkedin.com/in/apasra-bansuan-rn-mhss-cphq-3781b722/"),
    ("تجمع الطائف الصحي", "Vice President of Digital Affairs (حسب RocketReach)", "إدارة شؤون رقمية", "https://rocketreach.co/taif-health-cluster-tjm-ltyf-lshy-management_b76ece79c53db9b9"),
    ("مستشفيات القوات المسلحة بالطائف", "Head of Data Management (حسب ZoomInfo وRocketReach)", "إدارة بيانات", "https://rocketreach.co/taif-armed-forces-hospitals-it-department_b5c04e14f42e080b"),
    ("مستشفيات القوات المسلحة بالطائف", "Director of Quality", "إدارة جودة", "https://www.linkedin.com/in/mohammed-abu-shanab-b8b130a9/"),
    ("مستشفيات القوات المسلحة بالطائف", "Quality Assurance Analyst سابقاً (نمذجة بيانات في Excel)، وتشتغل الحين Senior Data Analyst في NEOM", "شغل بيانات داخل الجودة", "https://www.linkedin.com/in/nuha-alshanbari-61619396/"),
    ("مستشفى الملك عبدالعزيز التخصصي بالطائف", "Senior Quality Specialist (قاد اعتماد CBAHI)", "جودة ومؤشرات", "https://www.linkedin.com/in/saud-alsharif-398a971a3/"),
    ("أمانة محافظة الطائف", "E-Services and Applications Consultant (يستخدم Power BI وSQL)", "شغل بيانات وتطبيقات داخل الأمانة", "https://www.linkedin.com/in/mohammed-noury-64a527a1/"),
    ("امتياز العربية للاستشارات (الطائف)", "IT Infrastructure Manager", "فريق IT", "https://www.linkedin.com/in/al-mqdashi-abdulmughni-0380791b0/"),
    ("مجموعة بن شافي التعليمية", "Senior System Administrator", "فريق IT", "https://www.linkedin.com/in/ahmed-awad-16529526/"),
    ("Taif Univest", "Technical Support", "دعم تقني", "https://www.linkedin.com/in/ghada-mohammad-678796a2/"),
    ("مطاعم الحاشي البري", "مدير مالي (الطائف)", "إدارة مالية (تقارير)", "https://www.linkedin.com/in/mohamed-heshmat-1949a2195/"),
    ("مستشفى العدواني العام", "موظفين في المستشفى ظهروا في بحث الجودة", "—", "https://www.linkedin.com/in/suhail-malik-53160b31/"),
]
sheet("أدلة من حسابات الموظفين", ["الجهة", "المسمى أو الدور اللي ظهر", "وش يدل عليه", "رابط الملف العام"], [36, 70, 30, 60], EVID, link_cols={4})

# ------------------------------------------------ recall test sheet
RECALL = [
    ("شركة توت", "وُجدت (Sabbar، عن طريق Google)", "توت المصيف للتجارة: أخصائي تسويق وأمين صندوق في الطائف."),
    ("إعلان مهندس بيانات (صورة في X، بدون اسم)", "لم يوجد", "الإعلانات المنشورة كصور في X ما تنقرأ آلياً، ولا يمكن الوصول لـ X من هنا."),
    ("شركة السفياني لأنظمة الأمن والسلامة", "وُجدت الجهة فقط", "موجودة كمقاول في الطائف، بدون إعلانات نصية. على الأغلب تعلن بصور."),
    ("شركة ماتريكس التقنية", "وُجدت الجهة فقط", "مجموعة ماتركس (الطائف وجدة)، بدون إعلانات منسوخة."),
    ("شركة العرض المتقن", "وُجدت (أي وظيفة + وظيفتك علينا)", "إعلانات تقنية في الطائف 2021–2023 عن طريق LinkedIn."),
    ("شركة سراكو", "وُجدت (تليجرام + Sabbar + Tanqeeb)", "أكثر من 30 إعلاناً في الطائف."),
    ("مذاق الملكة", "وُجدت الجهة فقط", "صفحة LinkedIn ومصنع بتقييم +A، بدون إعلانات نصية."),
    ("تحكم", "وُجدت (أي وظيفة)", "إعلانات إدارية وتقنية في الرياض والطائف."),
    ("إمداد", "لم يوجد في الطائف", "إمداد الخبرات أعلنت في مدن أخرى. يحتاج الاسم الدقيق."),
    ("NHC (الوطنية للإسكان)", "إعلانات وطنية فقط", "ما فيه إعلان خاص بالطائف."),
    ("شركة العربي الطبية", "وُجدت (Tanqeeb / Sabbar)", "شركة العرب العصري الطبي: 8 إعلانات حالية في الطائف."),
    ("شركة إنصات", "لم توجد", "ما فيه أي أثر نصي بالعربي أو الإنجليزي. أرسل الرابط إذا عندك."),
    ("شركة عنب الطائف", "وُجدت (تليجرام)", "محاسب، أكتوبر 2026."),
    ("مجموعة تركي عمر السفياني", "وُجدت جزئياً (X)", "منشور في حساب (وظائف الطائف) على X: منظمين وإعلاميين ومصورين."),
    ("شركات الفنادق والسياحة", "وُجدت (تليجرام)", "إعلانات فنادق كثيرة في الطائف، أغلبها استقبال. فلتر القطاع (ضيافة وسياحة) في ورقة الجهات."),
    ("وكالة أركيزا ميديا", "لم توجد", "ما فيه أثر نصي. على الأغلب تعلن بالصور في Instagram."),
    ("شركة محمد الأصيفر", "وُجدت (تليجرام + saudijobs24)", "8 إعلانات من 2022 إلى 2026 (تسويق، محاسبة، موارد بشرية). كانت ضايعة بسبب ثغرة أصلحتها."),
]
sheet("اختبار التغطية (قائمتك)", ["الجهة اللي ذكرتها", "النتيجة", "التفاصيل"], [40, 34, 90], RECALL)

# ------------------------------------------------ channels + method + excluded
CH2 = CHANNELS + [
    ("تليجرام", "قناة وظائف الطائف 24 (أرشيف من 2020)", "https://t.me/s/jobstaif", "أرشيف كامل يمكن البحث فيه. سحبته كله في هذا الملف."),
    ("أي وظيفة", "بحث (الطائف) في أرشيف أي وظيفة", "https://www.ewdifh.com/search?word=%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81", "ينسخ إعلانات LinkedIn وتويتر كنص، ومفيد للإعلانات القديمة."),
    ("وظيفتك علينا", "أرشيف وظائف الطائف", "https://www.wdeftksa.com/sa/jobs/city-%D9%88%D8%B8%D8%A7%D8%A6%D9%81-%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81", "حوالي 548 صفحة أرشيف."),
    ("Tanqeeb", "وظائف الطائف الحالية (يجمع Bayt وSabbar وجدارات)", "https://saudi.tanqeeb.com/ar/s/%D9%88%D8%B8%D8%A7%D8%A6%D9%81/%D9%88%D8%B8%D8%A7%D8%A6%D9%81-%D9%81%D9%8A-%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81", ""),
    ("LinkedIn — شركات مقرها الطائف (Google)", "صفحات الشركات اللي مقرها الطائف", g('site:linkedin.com/company "Taif, Makkah"'), "غيّر الكلمات: manufacturing، hospital، trading، IT."),
]
sheet("قنوات البحث", ["القناة", "الغرض", "الرابط الجاهز", "طريقة الاستخدام"], [26, 40, 70, 60], CH2, link_cols={3})
sheet("مستبعدة أو ضعيفة", ["الجهة", "الاسم الإنجليزي", "سبب الاستبعاد", "المصدر"], [34, 30, 80, 60], EXCL, link_cols={4})

n_jobs = len(job_rows); n_emp = len(emp_rows)
n_data_jobs = sum(1 for r in job_rows if r[6].startswith('بيانات'))
n_dom = {n: sum(1 for r in job_rows if n in r[7]) for n in DOM_NAMES}; n_it_jobs = sum(1 for r in job_rows if r[6].startswith('تقنية'))
METHOD2 = [
    ("تاريخ التحديث", TODAY + " (الإصدار الرابع)"),
    ("ما الجديد في الإصدار الرابع", "1) وسّعت التصنيف ليغطي مجالات البيانات الست بالعربي والإنجليزي: تحليل وBI، وهندسة بيانات (ETL وELT ومستودعات وقواعد بيانات وSQL)، وعلم بيانات وذكاء اصطناعي، وحوكمة وجودة وإدارة بيانات (DMO وNDMO وPDPL والبيانات المرجعية)، وإحصاء ومعلوماتية صحية، ووثائق وسجلات. وصار لكل مجال عمود تفلتر عليه. 2) صححت مشكلة: الإصدار الثالث كان يحسب 50 إعلان بيانات، لكن جزءاً كبيراً منها ضوضاء، لأن إعلانات خدمة العملاء والتسويق فيها وصف قالبي مثل (بناء قاعدة بيانات للعملاء) و(مؤشرات الأداء). الحين التصنيف يعتمد على العنوان، ومن النص الكامل ما يأخذ إلا عبارات قوية مثل (مهندس بيانات) و(SQL) و(ETL) و(حوكمة البيانات). الرقم الحقيقي أقل، لكنه أدق. 3) أضفت عمود (مكتب إدارة البيانات NDMO وحماية البيانات) في ورقة الأولوية، ويفرّق بين الإلزامي والإرشادي. 4) أضفت إعلانات BAE للبيانات المرجعية والكوادر الوطنية، وإعلانات هندسة بيانات غير مؤكدة الموقع (Schneider Electric وJobmel)."),
    ("المرجع النظامي", "ضوابط ومواصفات إدارة البيانات الوطنية وحوكمتها وحماية البيانات الشخصية (مكتب إدارة البيانات الوطنية، سدايا): إلزامية على الجهات الحكومية، وتنص على تعيين مسؤول لإدارة البيانات وقياس الامتثال سنوياً. والأمر السامي رقم 59766 نص على إنشاء مكاتب بيانات في الجهات الحكومية مرتبطة بالمسؤول الأول. ونظام حماية البيانات الشخصية إلزامي على كل من يعالج بيانات شخصية في المملكة، حكومي أو خاص. ملاحظة: تطبيق الضوابط على مقاولي القطاع الخاص اللي يديرون بيانات حكومية مصدره ملخصات شركات استشارية، فتحقق منه من النص الرسمي: https://sdaia.gov.sa/ndmo/Files/Policies001.pdf"),
    ("الإصدار الثالث", "1) أصلحت ثغرة: منشورات قناة التليجرام اللي ما تذكر كلمة (الطائف) كانت مستبعدة. الحين أتحقق من مدينة كل إعلان من نصه الكامل في saudijobs24، وصار التصنيف أدق لأنه يقرأ الإعلان كامل مو بس العنوان. 2) ورقة جديدة: أدلة من حسابات الموظفين في LinkedIn. 3) أضفت الشركات اللي ذكرتها، وصف لقطاع الفنادق والسياحة. 4) ما قدرت أصل لجدارات وأرشيف الإنترنت (Wayback) لأن البيئة رفضت إضافتهم."),
    ("الإصدار الثاني", f"1) سحب آلي لأرشيف قناة (وظائف الطائف 24) في تليجرام، حوالي 15 ألف منشور من 2020 إلى 2026. 2) سحب أرشيف (وظيفتك علينا)، 548 صفحة. 3) سحب نتائج (الطائف) في أرشيف (أي وظيفة)، 439 إعلاناً، وهذا الموقع ينسخ إعلانات LinkedIn وتويتر كنص. 4) وظائف الطائف الحالية من Tanqeeb، وهو يجمع Bayt وSabbar وجدارات. 5) فحص صفحات حوالي 1300 شركة في Sabbar. 6) بحث في Google عن صفحات الشركات في LinkedIn اللي مقرها الطائف. النتيجة: {n_emp} جهة و{n_jobs} إعلاناً، منها {n_data_jobs} إعلان بيانات/تحليل و{n_it_jobs} إعلان تقنية معلومات."),
    ("التصنيف الآلي", "كل إعلان صُنّف آلياً حسب كلمات مفتاحية: (بيانات/تحليل/إحصاء) لكلمات مثل محلل، إحصائي، Power BI، SQL، أرشفة، معلوماتية. و(تقنية معلومات) لكلمات مثل IT، شبكات، مطور، أمن سيبراني. وإعلانات إدخال البيانات معلّمة (مستبعد). التصنيف الآلي يخطئ أحياناً، فراجع العنوان والرابط."),
    ("حالي / سابق", "أي إعلان من سبتمبر 2026 وبعد مكتوب (حالي/حديث)، لكن تحقق من الرابط لأن بعض الإعلانات تُغلق بسرعة."),
    ("قيد — LinkedIn وX", "صفحات LinkedIn ومنشوراتها وموقع X ما تُفتح آلياً من هنا، ومحركات البحث ما تفهرس منشورات LinkedIn جيداً. الجزء اللي غطيته منها هو اللي نسخه موقعا (أي وظيفة) و(وظيفتك علينا) كنص."),
    ("قيد — إعلانات الصور", "الإعلانات المنشورة كصور فقط (مثل إعلان مهندس البيانات في X) ما يمكن قراءتها آلياً إلا إذا نسخها أحد كنص. لهذا بعض الجهات في قائمتك ظهرت كجهة بدون إعلانات."),
    ("قيد — أسماء الجهات", "أسماء الجهات استُخرجت آلياً من نص الإعلانات، فقد تتكرر الجهة الواحدة باسمين أو يظهر اسم ناقص. استخدم الفلترة والبحث في Excel."),
    ("قيد — البريد والهاتف", "ما كتبت أي إيميل أو رقم ما تحققت منه. تخمين إيميل غلط يضيّع عليك فرصة."),
    ("قيد — الشمول", "حتى بعد كل هذا، القائمة ما تقدر تكون كاملة 100%، لأن جزءاً من الإعلانات موجود فقط في واتساب أو صور في X أو منشورات LinkedIn. ورقة (اختبار التغطية) توضح النسبة على قائمتك."),
] + [m for m in METHOD if m[0] in ("درجات الثقة", "الأولوية", "قيد 3 — مصادر تجميع")]
sheet("المنهجية والقيود", ["البند", "التفاصيل"], [24, 140], METHOD2)

# ------------------------------------------------ summary
sm = wb.active
sm.title = "ملخص"
sm.sheet_view.rightToLeft = True
sm["A1"] = "جهات ووظائف البيانات في الطائف — الإصدار الرابع"; sm["A1"].font = Font(bold=True, size=14, name="Arial")
sm["A2"] = f"تاريخ التحديث: {TODAY}"; sm["A2"].font = BODY_FONT
lines = [("الجهات ذات الأولوية (مختارة يدوياً بأدلة)", len(R)), ("شركات الطائف للتقديم اليدوي (خاص)", N_MANUAL), ("كل الجهات اللي أعلنت في الطائف", n_emp),
         ("كل الإعلانات (حالية وسابقة)", n_jobs), ("إعلانات بيانات (كل المجالات)", n_data_jobs)] + [("   منها: " + k, v) for k, v in n_dom.items()] + [("إعلانات تقنية معلومات", n_it_jobs)]
sm["A4"], sm["B4"] = "البند", "العدد"
for c in ("A4", "B4"):
    sm[c].fill, sm[c].font, sm[c].alignment = HDR_FILL, HDR_FONT, CENTER
for i, (a, b) in enumerate(lines, 5):
    sm[f"A{i}"], sm[f"B{i}"] = a, b
    sm[f"A{i}"].font = sm[f"B{i}"].font = BODY_FONT
sm["D4"], sm["E4"] = "حالة التقديم (ورقة الأولوية)", "العدد"
for c in ("D4", "E4"):
    sm[c].fill, sm[c].font, sm[c].alignment = HDR_FILL, HDR_FONT, CENTER
stl = get_column_letter(keys.index("_status") + 1)
for i, t in enumerate(["لم أقدّم", "قدّمت إلكترونياً", "قدّمت يدوياً", "تواصلت مع شخص", "مقابلة", "عرض", "مرفوض", "مؤجل"], 5):
    sm[f"D{i}"] = t
    sm[f"E{i}"] = f"=COUNTIF('الجهات ذات الأولوية'!{stl}:{stl},\"{t}\")"
sm.column_dimensions["A"].width = 44; sm.column_dimensions["B"].width = 10; sm.column_dimensions["C"].width = 3
sm.column_dimensions["D"].width = 28; sm.column_dimensions["E"].width = 10
sm["A20"] = ("طريقة الاستخدام:\n1) ابدأ بورقة (الجهات ذات الأولوية) وفلتر الأولوية (أ).\n2) ورقة (الوظائف) فيها كل إعلان مع رابطه. فلتر عمود (التصنيف) على (بيانات) أو (تقنية).\n"
             "3) ورقة (شركات الطائف للتقديم اليدوي) فيها شركات خاصة قائمة في الطائف، تعلن أو ما تعلن، ومع كل وحدة الدور اللي تعرضه عليها.\n4) ورقة (كل الجهات المعلنة) فيها كل الجهات اللي أعلنت في الطائف. رتّبها حسب عدد الإعلانات عشان تكتشف الجهات الكبيرة.\n5) سجّل حالة تقديمك من القائمة المنسدلة.")
sm["A20"].alignment = WRAP; sm["A20"].font = BODY_FONT
sm.merge_cells("A20:E27")
ms = wb['شركات الطائف للتقديم اليدوي']
wb.move_sheet(ms, offset=-(wb.sheetnames.index(ms.title) - 2))
wb.save(OUT)
print("saved", OUT, "curated", len(R), "employers", n_emp, "jobs", n_jobs, "data", n_data_jobs, "it", n_it_jobs)
