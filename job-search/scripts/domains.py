import re
# Data domains (DAMA-aligned). Arabic + English. Order = display order.
DOMAINS = [
 ("تحليل البيانات وذكاء الأعمال", r"تحليل البيانات|تحليل بيانات|محلل بيانات|محللة بيانات|محلل أعمال|محلل اعمال|ذكاء الأعمال|ذكاء اعمال|لوحات المؤشرات|لوحات معلومات|مؤشرات الأداء|مؤشرات الاداء|أخصائي تقارير|اخصائي تقارير|إعداد التقارير التحليلية|Power ?BI|Tableau|Qlik|Looker|business intelligence|\bBI\b|data analys|analytics|\bKPIs?\b|dashboard|reporting analyst|\bMIS\b"),
 ("هندسة البيانات وقواعد البيانات", r"هندسة البيانات|مهندس بيانات|مهندس البيانات|مستودع البيانات|مستودعات البيانات|بحيرة البيانات|تكامل البيانات|نقل البيانات|قواعد البيانات|قاعدة بيانات|قواعد بيانات|مدير قواعد|مسؤول قواعد|\bETL\b|\bELT\b|data engineer|data warehouse|data lake|data pipeline|data integration|\bSQL\b|PL/SQL|Oracle Database|SQL Server|PostgreSQL|MySQL|MongoDB|\bDBA\b|database|Informatica|Talend|\bSSIS\b|\bAirflow\b|\bSpark\b|\bHadoop\b|Databricks|Snowflake|BigQuery"),
 ("علم البيانات والذكاء الاصطناعي", r"علم البيانات|عالم بيانات|علوم البيانات|ذكاء اصطناعي|الذكاء الاصطناعي|تعلم الآلة|تعلم الاله|التعلم الآلي|التعلم العميق|نمذجة تنبؤية|data scien|machine learning|deep learning|\bML\b|\bAI\b|artificial intelligence|predictive|\bNLP\b|computer vision|Python|\bR\b programming|TensorFlow|PyTorch|scikit"),
 ("حوكمة وإدارة وجودة البيانات (DMO)", r"حوكمة البيانات|إدارة البيانات|ادارة البيانات|مكتب إدارة البيانات|مكتب ادارة البيانات|جودة البيانات|سلامة البيانات|البيانات المرجعية|البيانات الرئيسية|البيانات الوصفية|أمين البيانات|امين البيانات|مالك البيانات|تصنيف البيانات|البيانات المفتوحة|حماية البيانات الشخصية|خصوصية البيانات|نظام حماية البيانات|مكتب إدارة البيانات الوطنية|سدايا|data governance|data management|data quality|data integrity|business data|national data|data specialist|data steward|data owner|master data|\bMDM\b|reference data|metadata|data catalog|data classification|open data|\bPDPL\b|\bNDMO\b|\bCDO\b|chief data|privacy"),
 ("الإحصاء والمعلوماتية الصحية", r"إحصائي|احصائي|إحصائية|الإحصاء الحيوي|احصاء حيوي|باحث إحصائي|المعلوماتية الصحية|معلوماتية صحية|المعلومات الصحية|معلومات صحية|ترميز طبي|الترميز الطبي|statistic|biostatistic|health informatics|informatics|medical coding|\bHIM\b"),
 ("إدارة الوثائق والسجلات", r"إدارة الوثائق|ادارة الوثائق|مراقب وثائق|مراقبة الوثائق|التحكم بالوثائق|الأرشفة الإلكترونية|الارشفة الالكترونية|خبير في الأرشفة|أخصائي أرشفة|إدارة السجلات|records management|document control|document controller|\bEDMS\b|archiv|فهرسة|الأرشيف|أرشيف"),
]
DOM_RX = [(n, re.compile(p, re.I)) for n, p in DOMAINS]
DE_RX = re.compile(r"إدخال بيانات|إدخال البيانات|ادخال بيانات|ادخال البيانات|مدخل بيانات|مدخلي بيانات|مدخلين بيانات|مدخل/ة بيانات|مدخلة بيانات|data entry", re.I)
IT_RX = re.compile(r"تقنية المعلومات|تقنية معلومات|نظم المعلومات|نظم معلومات|\bIT\b|برمج|مطور|developer|\bERP\b|\bSAP\b|Odoo|أمن سيبراني|الأمن السيبراني|cyber|حاسب آلي|حاسب الي|علوم الحاسب|computer|شبكات|\bnetwork|system admin|systems? analyst|محلل نظم|دعم فني تقني|دعم تقني|help ?desk", re.I)

def domains_of(text):
    t = text or ''
    t2 = DE_RX.sub(' ', t)
    return [n for n, rx in DOM_RX if rx.search(t2)]

# Strict patterns for long ad bodies (templated descriptions mention KPIs/databases generically)
STRICT = {
 "تحليل البيانات وذكاء الأعمال": r"(?:وظيفة|بمسمى|مسمى|مطلوب|المسمى الوظيفي[:\s–-]*)\s*(?:محلل بيانات|محللة بيانات|محلل أعمال|أخصائي تحليل بيانات|أخصائي ذكاء أعمال)|Power ?BI|Tableau|data analyst|BI analyst|BI developer|business intelligence",
 "هندسة البيانات وقواعد البيانات": r"مهندس بيانات|مهندس البيانات|مطور قواعد بيانات|مطور قواعد البيانات|مسؤول قواعد البيانات|مدير قواعد البيانات|\bETL\b|\bELT\b|data engineer|data warehouse|\bDBA\b|database administrator|database developer",
 "علم البيانات والذكاء الاصطناعي": r"عالم بيانات|أخصائي علم البيانات|مهندس ذكاء اصطناعي|مهندس تعلم الآلة|machine learning engineer|data scientist",
 "حوكمة وإدارة وجودة البيانات (DMO)": r"أخصائي حوكمة البيانات|حوكمة البيانات|أخصائي جودة البيانات|جودة البيانات|مكتب إدارة البيانات|حماية البيانات الشخصية|data governance|data quality|master data|data steward|\bNDMO\b|\bPDPL\b",
 "الإحصاء والمعلوماتية الصحية": r"(?:بمسمى|مسمى|وظيفة|المسمى الوظيفي[:\s–-]*)\s*(?:إحصائي|احصائي|أخصائي إحصاء|اخصائي احصاء|أخصائي معلوماتية صحية)|معلوماتية صحية|ترميز طبي|health informatics|statistician",
 "إدارة الوثائق والسجلات": r"مراقب وثائق|مراقب الوثائق|document controller|أخصائي أرشفة|خبير في الأرشفة|records management",
}
STRICT_RX = [(n, re.compile(p, re.I)) for n, p in STRICT.items()]

def domains_of_ad(title, body=''):
    t = DE_RX.sub(' ', title or '')
    b = DE_RX.sub(' ', body or '')
    out = [n for n, rx in DOM_RX if rx.search(t)]
    for n, rx in STRICT_RX:
        if n not in out and rx.search(b):
            out.append(n)
    order = [d[0] for d in DOMAINS]
    return sorted(out, key=order.index)
