import sys
from urllib.parse import quote_plus
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

OUT = sys.argv[1]
TODAY = "2026-11-18"

# ---------------------------------------------------------------- columns
COLS = [
    ("م", 5),
    ("الأولوية", 9),
    ("اسم الجهة (عربي)", 30),
    ("اسم الجهة (إنجليزي)", 30),
    ("القطاع / المجال", 18),
    ("نوع الجهة", 14),
    ("الحجم التقديري", 16),
    ("الموقع في الطائف", 26),
    ("هل فيها قسم/وظائف بيانات؟", 16),
    ("نوع الدليل", 22),
    ("تفاصيل الدليل", 60),
    ("تاريخ الدليل", 14),
    ("مسميات البيانات اللي ظهرت", 34),
    ("القسم المتوقع للتقديم", 26),
    ("ملاءمة لمستواك (حديث تخرج)", 16),
    ("شروط مهمة / ملاحظات", 40),
    ("طريقة التقديم", 34),
    ("رابط التوظيف", 40),
    ("الموقع الإلكتروني", 32),
    ("LinkedIn", 38),
    ("X / تويتر", 26),
    ("البريد / الهاتف", 24),
    ("شخص للتواصل (من مصدر عام)", 34),
    ("كيف وصلت لها", 46),
    ("المصدر 1", 40),
    ("المصدر 2", 40),
    ("المصدر 3", 40),
    ("الخطوة التالية المقترحة", 44),
    ("حالة التقديم", 14),
    ("تاريخ التقديم", 13),
    ("ملاحظاتي", 30),
]

NA = "لم أتحقق منه — ابحث في الموقع الرسمي"

# each row: dict keyed by column header
R = []
def add(**k):
    R.append(k)

add(p="أ", ar="تجمع الطائف الصحي (ومستشفياته التابعة)", en="Taif Health Cluster",
    sector="صحة", typ="حكومي / شبه حكومي (الصحة القابضة)", size="كبير جداً (آلاف الموظفين عبر المستشفيات)",
    loc="الطائف — الإدارة العامة للتجمع + مستشفيات التجمع",
    has="مؤكد", ev="موظف حالي على LinkedIn + هيكل إداري",
    evd="ملف LinkedIn عام لموظفة بمسمى Associate Executive Director of Business Intelligence في التجمع منذ أبريل 2024، وسبق لها مسمى Data Specialist داخل التجمع. ويذكر RocketReach وجود منصب Vice President of Digital Affairs. أي إن عند التجمع إدارة ذكاء أعمال (BI) وإدارة شؤون رقمية.",
    date="2024 – حالياً",
    roles="Business Intelligence، Data Specialist، Digital Affairs، المعلوماتية الصحية",
    dept="إدارة ذكاء الأعمال (BI) / الشؤون الرقمية / المعلوماتية الصحية / الجودة والأداء",
    fit="عالية",
    cond="التوظيف عبر جولات تعلنها الصحة القابضة، ومدة التسجيل قصيرة (مثلاً 15 يوماً من 3 مارس 2026). المسميات الإدارية والتقنية قد تأتي بأسماء غير صريحة مثل (أخصائي معلوماتية صحية) أو (أخصائي إحصاء).",
    how="بوابة توظيف التجمع + بوابة الصحة القابضة، ومعها تواصل مباشر مع إدارة BI (تقديم يدوي في مقر الإدارة العامة للتجمع)",
    jobs="https://altaif-careers.health.sa", web="https://www.health.sa/ar/careers",
    li=NA, x="https://x.com/TaifHCluster", mail=NA,
    contact="Associate Executive Director of Business Intelligence — https://www.linkedin.com/in/hanan-a-b70296160/",
    found="بحثت عن (Taif Health Cluster + Business Intelligence / Data Analyst) داخل LinkedIn عن طريق محرك البحث، فطلع ملف موظفة قيادية في BI. وأكدت وجود الإدارة الرقمية من RocketReach.",
    s1="https://www.linkedin.com/in/hanan-a-b70296160/",
    s2="https://rocketreach.co/taif-health-cluster-tjm-ltyf-lshy-management_b76ece79c53db9b9",
    s3="https://x.com/TaifHCluster/status/2028483134543565138",
    nxt="1) فعّل تنبيهات حساب X للتجمع. 2) أرسل رسالة قصيرة لمديرة BI على LinkedIn تعرّف فيها بنفسك وتدريبك في مكتب إدارة البيانات بالجامعة. 3) جهّز ملفك مسبقاً لأن جولات التوظيف قصيرة.")

add(p="أ", ar="مستشفيات القوات المسلحة بالطائف (الهدا – الأمير منصور – الأمير سلطان العسكري)", en="Taif Armed Forces Hospitals (Al Hada / Prince Mansour)",
    sector="صحة عسكرية", typ="حكومي (الخدمات الطبية للقوات المسلحة – وزارة الدفاع)", size="كبير جداً",
    loc="الهدا + الطائف",
    has="مؤكد", ev="هيكل وظيفي (منصب قيادي) + موظفين سابقين",
    evd="ZoomInfo وRocketReach يذكران منصب Head of Data Management في المستشفيات، وقسم IT فيه حوالي 34 موظفاً (بيانات من جهات تجميع، فتحقق منها). وفيه ملف LinkedIn لمحللة بيانات حالية في NEOM عملت سابقاً Quality Assurance Analyst في المستشفيات (2018) ومن مهامها نمذجة البيانات في Excel.",
    date="2018 – حالياً",
    roles="Data Management، Medical Informatics، Statistic Technician (فني إحصاء)، Quality Analyst",
    dept="إدارة البيانات / تقنية المعلومات / الجودة / الإحصاء الطبي",
    fit="عالية",
    cond="إعلانات المسميات المشابهة في مستشفيات القوات المسلحة الأخرى (مثل فني إحصاء في نجران وأخصائي معلوماتية في الرياض) تنشر على بوابة الخدمات الطبية، فراقب البوابة واختر الطائف.",
    how="بوابة توظيف الخدمات الطبية للقوات المسلحة، ومعها زيارة شؤون الموظفين وإدارة البيانات",
    jobs="https://talents.msd.med.sa", web=NA,
    li="https://www.linkedin.com/company/al-hada-hospital", x=NA, mail=NA,
    contact="Head of Data Management (الاسم في ZoomInfo/RocketReach) — https://rocketreach.co/taif-armed-forces-hospitals-it-department_b5c04e14f42e080b",
    found="بحثت عن (Taif Armed Forces Hospitals + data analyst / health informatics / BI). طلعت صفحات ZoomInfo وRocketReach فيها منصب رئيس إدارة البيانات، وطلع ملف موظفة سابقة.",
    s1="https://www.zoominfo.com/pic/taif-armed-forces-hospitals/427201675",
    s2="https://rocketreach.co/taif-armed-forces-hospitals-it-department_b5c04e14f42e080b",
    s3="https://www.linkedin.com/in/nuha-alshanbari-61619396/",
    nxt="ابحث في LinkedIn عن موظفي (Taif Armed Forces Hospitals) بمسمى Data أو Informatics، وتواصل مع رئيس إدارة البيانات. وراقب talents.msd.med.sa.")

add(p="أ", ar="جامعة الطائف", en="Taif University",
    sector="تعليم عالي", typ="حكومي", size="كبير جداً",
    loc="الحوية – الطائف",
    has="مؤكد", ev="معرفة مباشرة (تدريبك) + برنامج ماجستير علم بيانات",
    evd="أنت تتدرب في مكتب إدارة البيانات بالجامعة، وهذا يكفي دليلاً. والجامعة تقدّم ماجستير في علم البيانات (2023–2025) حسب ملف خريجة تعمل في BI بتجمع الطائف الصحي، أي إن الجامعة نفسها تخرّج كوادر بيانات للسوق المحلي.",
    date="حالياً",
    roles="مكتب إدارة البيانات (DMO)، تقنية المعلومات، التحول الرقمي، القياس والتقويم",
    dept="مكتب إدارة البيانات / عمادة تقنية المعلومات / وحدة القياس والإحصاء",
    fit="عالية جداً (أنت معروف عندهم)",
    cond="التعيين الرسمي غالباً عن طريق جدارات أو بعقود. والأقوى لك هو التحويل من متدرب إلى متعاقد.",
    how="داخلي عن طريق مشرفك في المكتب + جدارات + عقود المشاريع",
    jobs="https://jadarat.sa", web="https://www.tu.edu.sa",
    li=NA, x=NA, mail=NA,
    contact="مشرفك في مكتب إدارة البيانات",
    found="من معلوماتك الشخصية، وتأكيد إضافي من ملف LinkedIn لخريجة ماجستير علم البيانات من الجامعة.",
    s1="https://www.tu.edu.sa",
    s2="https://www.linkedin.com/in/hanan-a-b70296160/",
    s3="",
    nxt="اسأل مشرفك مباشرة: هل فيه احتمال عقد أو وظيفة بعد التدريب؟ ومن الجهات في الطائف اللي يعرف إن عندها احتياج بيانات؟ (شبكة علاقاته مصدر قوي).")

add(p="أ", ar="هيئة تطوير محافظة الطائف", en="Taif Development Authority",
    sector="تطوير حضري / تخطيط", typ="حكومي (أُنشئت في 5 مايو 2022)", size="متوسط وفي مرحلة بناء الفريق",
    loc="الطائف (تحقق من المقر، فبعض المصادر تذكر مكاتب خارجها)",
    has="مؤكد", ev="إعلان سابق (2026)",
    evd="في يونيو 2026 أعلنت الهيئة عبر جدارات عن 5 وظائف قيادية، منها (مدير إدارة البيانات)، والتخصصات المقبولة تشمل إدارة البيانات وعلم البيانات وتحليل البيانات وعلوم الحاسب ونظم المعلومات. وجود إدارة بيانات جديدة يعني على الأغلب احتياجاً لاحقاً لأخصائيين تحت المدير.",
    date="يونيو 2026 (آخر موعد 22 يونيو 2026)",
    roles="مدير إدارة البيانات (وأدوار أخصائيين متوقعة لاحقاً)",
    dept="إدارة البيانات / التخطيط الاستراتيجي / التميز المؤسسي",
    fit="متوسطة الآن وعالية مستقبلاً",
    cond="الإعلان كان لوظائف قيادية، فلا يناسبك مباشرة، لكنه مؤشر قوي على بناء فريق بيانات.",
    how="جدارات + متابعة الهيئة على LinkedIn وX + تواصل مع مدير إدارة البيانات بعد تعيينه",
    jobs="https://jadarat.sa", web=NA,
    li=NA, x=NA, mail=NA,
    contact="مدير إدارة البيانات (ابحث عنه في LinkedIn بعد التعيين)",
    found="بحثت عن (أمانة الطائف + أخصائي بيانات)، فظهر ضمن النتائج إعلان هيئة تطوير الطائف لعام 2026، وبعدها تحققت منه بعدة مصادر.",
    s1="https://www.saudijobs24.com/t180709-topic",
    s2="https://www.forsawazifasa.com/2026/06/taif-jobs.html",
    s3="https://saudipedia.com/en/article/2584/government-and-politics/commissions/taif-development-authority",
    nxt="ابحث في LinkedIn عن (Taif Development Authority) و(هيئة تطوير الطائف)، وتابع من عُيّن مديراً للبيانات وتواصل معه مبكراً.")

add(p="أ", ar="أمانة محافظة الطائف", en="Taif Municipality (Amanah)",
    sector="بلديات", typ="حكومي", size="كبير",
    loc="الطائف + البلديات الفرعية",
    has="محتمل قوي", ev="لائحة تنظيمية إلزامية + إعلانات سابقة في جدارات",
    evd="لائحة وزارة البلديات والإسكان (قرار 4200018300 بتاريخ 7/1/1442هـ) تنظّم وحدات البيانات البلدية والإحصاء في الأمانات، والوزارة تقيس أداء هذه الوحدات دورياً. والأمانة أعلنت سابقاً في جدارات وظائف تقنية مثل (أخصائي أمن معلومات) في 2024.",
    date="1442هـ (اللائحة) / 2024 (إعلانات)",
    roles="وحدة البيانات والإحصاء، أمن المعلومات، التحول الرقمي",
    dept="وحدة البيانات البلدية والإحصاء / تقنية المعلومات",
    fit="عالية",
    cond="التوظيف الرسمي عبر جدارات بمراتب (الخامسة إلى الثامنة). وفيه أحياناً وظائف بالتعاقد.",
    how="جدارات + زيارة إدارة الموارد البشرية ووحدة البيانات",
    jobs="https://jadarat.sa", web="https://my.gov.sa/ar/agencies/17261",
    li=NA, x=NA, mail="الإيميل والهاتف في صفحة بلدي: https://www.balady.gov.sa/en/node/11793",
    contact="",
    found="ما لقيت إعلان (أخصائي بيانات) صريح، لكن لقيت لائحة الوزارة اللي تُلزم الأمانات بوحدات بيانات وإحصاء، ولقيت إعلانات الأمانة السابقة في جدارات.",
    s1="https://momah.gov.sa/sites/default/files/2025-10/allayht%20altnzymyt%20lwhdat%20albyanat%20albldyt%20walahsa%20fy%20alqta%60%20albldy.pdf",
    s2="https://ajel.sa/thehour/3r9wvh5uti",
    s3="https://www.balady.gov.sa/en/node/11793",
    nxt="اسأل في الأمانة عن (وحدة البيانات والإحصاء)، وتعرّف على رئيسها. خبرتك في إطار DAMA ومعايير NDMO نقطة قوية عندهم.")

add(p="ب", ar="بي إيه إي سيستمز السعودية — قاعدة الملك فهد الجوية (مركز دعم التايفون)", en="BAE Systems Saudi Arabia — Taif (Typhoon Support Centre)",
    sector="دفاع وطيران", typ="خاص (مقاول دفاعي)", size="كبير جداً",
    loc="قاعدة الملك فهد الجوية – الطائف",
    has="مؤكد", ev="إعلانات حالية وسابقة متعددة",
    evd="عدة إعلانات في الطائف تحديداً: Data Management Technician (Level 1)، وIn‑Service Data Analysis (ISDA) Specialist بأدوات Oracle SQL وExcel، وSenior Technical Specialist – Data Integrity (للسعوديين)، وESS Business Data Senior Specialist (قائد فريق سلامة بيانات الأعمال BDIT)، وSenior Data Analyst Specialist (نُشر 3 فبراير 2025).",
    date="2025 – 2026",
    roles="Data Management Technician، ISDA Specialist، Data Integrity، Business Data Specialist، Senior Data Analyst",
    dept="Data Management Cell / Business Data Integrity Team",
    fit="متوسطة (أغلبها يطلب خبرة 5–10 سنوات، والأقرب لك Level 1)",
    cond="العمل داخل قاعدة عسكرية، وقد تُطلب موافقات أمنية. بعض الوظائف للسعوديين فقط، ومعظمها يطلب خبرة.",
    how="بوابة BAE الرسمية بفلتر Taif. والتقديم اليدوي غير عملي لأن المقر داخل القاعدة.",
    jobs="https://jobsearch.baesystems.com", web="https://jobsearch.baesystems.com",
    li=NA, x=NA, mail=NA,
    contact="",
    found="بحثت عن (site:linkedin.com/jobs Taif data)، فظهر إعلان Data Management Technician على Sabbar. بعدها بحثت عن BAE Taif data وطلعت بقية الإعلانات.",
    s1="https://jobsearch.baesystems.com/job/data-management-technician-level-1-00132782",
    s2="https://jobsearch.baesystems.com/job/in-service-data-analysis-isda-specialist-00134354",
    s3="https://jobs.y-axis.com/jobs/senior-data-analyst-specialist-bae-systems",
    nxt="أنشئ تنبيه وظائف في بوابة BAE (Location: Taif, keyword: data)، وقدّم على أي وظيفة Level 1 أو Technician.")

add(p="ب", ar="مدينة الملك سلمان التخصصي للحرس الوطني بالطائف (الشؤون الصحية بالحرس الوطني)", en="King Salman Specialized Hospital – MNGHA Taif",
    sector="صحة", typ="حكومي (وزارة الحرس الوطني)", size="كبير",
    loc="الطائف",
    has="محتمل", ev="استنتاج من هيكل الجهة الأم",
    evd="الشؤون الصحية بالحرس الوطني عندها أقسام معلوماتية صحية وبيانات على مستوى المملكة، وبوابة التوظيف فيها فلتر لمدينة الطائف. لكن لقطة من أبريل 2025 أظهرت 0 وظائف في الطائف وقتها.",
    date="أبريل 2025 (لقطة البوابة)",
    roles="Health Informatics، Medical Records، Decision Support (متوقع)",
    dept="المعلوماتية الصحية / السجلات الطبية / الجودة",
    fit="متوسطة",
    cond="",
    how="بوابة توظيف الحرس الوطني الصحية بفلتر الطائف",
    jobs="https://careers.mngha.med.sa", web="https://mngha.med.sa/english/digital-channels/Pages/careers.aspx",
    li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن مستشفيات الطائف الكبيرة، وبعدها بحثت عن (King Salman Specialized Hospital Taif + informatics).",
    s1="https://careers.mngha.med.sa/",
    s2="https://www.zavis.ai/ar/sa/directory/taif/hospitals/king-salman-specialized-hospital-of-the-national-guard",
    s3="",
    nxt="راجع البوابة شهرياً، وابحث في LinkedIn عن موظفي المستشفى في IT أو Informatics.")

add(p="ب", ar="شركة المياه الوطنية — الطائف", en="National Water Company (NWC) – Taif",
    sector="مرافق / مياه", typ="شبه حكومي", size="كبير جداً (على مستوى المملكة)",
    loc="الطائف",
    has="محتمل قوي", ev="إعلانات سابقة",
    evd="أعلنت الشركة في برنامج تمهير وظائف في الرياض والطائف، منها (محلل نظم تقنية المعلومات). وفي 15 يونيو 2026 نُشر إعلان في الطائف لموظفة تحليل بيانات عند شركة تعمل بالتعاون مع المياه الوطنية (انظر الصف التالي).",
    date="إعلان تمهير (قديم) + يونيو 2026",
    roles="محلل نظم تقنية المعلومات، تحليل بيانات وتقارير",
    dept="تقنية المعلومات / التقارير التشغيلية / خدمة العملاء",
    fit="عالية (برامج تمهير لحديثي التخرج)",
    cond="بعض وظائف الطائف في إعلان تمهير كانت للرجال فقط وبعضها للجنسين.",
    how="بوابة توظيف الشركة + منصات تمهير وطاقات",
    jobs=NA, web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن (وظائف الطائف + تحليل بيانات)، فظهر إعلان المقاول المرتبط بالمياه الوطنية، وبعدها بحثت عن إعلانات الشركة نفسها في الطائف.",
    s1="https://www.ewdifh.com/jobs/56388",
    s2="https://www.jobs-arab.com/sa/jobs/1149772002/",
    s3="",
    nxt="ابحث في LinkedIn عن (National Water Company Taif) بمسمى Data أو Reporting، واسأل عن مقاولي التشغيل في الطائف.")

add(p="ب", ar="شركة مقاولة/تشغيل (غير مسمّاة في الإعلان) تعمل مع المياه الوطنية بالطائف", en="(Unnamed contractor working with NWC – Taif)",
    sector="خدمات تشغيل", typ="خاص", size="غير معروف",
    loc="الطائف",
    has="مؤكد", ev="إعلان سابق",
    evd="إعلان بتاريخ 15 يونيو 2026: مطلوب موظفة سعودية، بكالوريوس، تتقن Excel وWord وPowerPoint ولديها القدرة على تحليل البيانات، للعمل في الطائف.",
    date="15 يونيو 2026",
    roles="موظفة تحليل بيانات",
    dept="",
    fit="عالية (لكن الإعلان كان للنساء)",
    cond="الإعلان كان مخصصاً للنساء، لكنه يدل على وجود احتياج تحليل بيانات عند مقاولي التشغيل في الطائف.",
    how="افتح الإعلان لمعرفة اسم الشركة وطريقة التواصل",
    jobs="https://www.jobs-arab.com/sa/jobs/1149772002/", web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت بالعربي عن (وظيفة محلل بيانات الطائف).",
    s1="https://www.jobs-arab.com/sa/jobs/1149772002/", s2="", s3="",
    nxt="افتح الإعلان وسجّل اسم الشركة، فإذا كان عندها مقر في الطائف قدّم يدوياً.")

add(p="ب", ar="الإدارة العامة للتعليم بمحافظة الطائف", en="Taif Education Department (MoE)",
    sector="تعليم عام", typ="حكومي", size="كبير جداً",
    loc="الطائف",
    has="محتمل قوي", ev="هيكل تنظيمي رسمي",
    evd="وزارة التعليم دمجت وكالتي التخطيط في (وكالة التخطيط والمعلومات). وفي كل إدارة تعليم إدارة تخطيط فيها وحدة (المعلومات وقياس الأداء).",
    date="هيكل حالي",
    roles="أخصائي معلومات، قياس أداء، إحصاء",
    dept="إدارة التخطيط — وحدة المعلومات وقياس الأداء",
    fit="متوسطة",
    cond="الوظائف المدنية فيها قليلة وتمر عبر جدارات.",
    how="جدارات + زيارة إدارة التخطيط والمعلومات",
    jobs="https://jadarat.sa", web=NA, li=NA, x=NA, mail=NA, contact="",
    found="استنتاج من هيكل وزارة التعليم الرسمي.",
    s1="https://moe.gov.sa/ar/updates-and-media/updates/planningmerge", s2="", s3="",
    nxt="زر إدارة التخطيط والمعلومات واسأل عن التعاقد أو المشاريع (مثل لوحات المؤشرات).")

add(p="ب", ar="مستشفى العدواني العام", en="Al Adwani General Hospital",
    sector="صحة", typ="خاص", size="متوسط (80 سريراً تشغيلياً و100 مرخّصاً)",
    loc="الطائف",
    has="محتمل", ev="قسم تقنية معلومات معلن في الموقع",
    evd="موقع المستشفى فيه صفحة لقسم تقنية المعلومات، وصفحة توظيف تطلب كوادر إدارية وداعمة. المستشفيات الخاصة تحتاج غالباً تقارير وجودة وترميز طبي وفوترة تأمين.",
    date="حالياً",
    roles="IT، تقارير، جودة، سجلات طبية (متوقع)",
    dept="تقنية المعلومات / الجودة / الإدارة المالية والتأمين",
    fit="عالية (التقديم اليدوي ممكن)",
    cond="",
    how="صفحة التوظيف + تقديم يدوي في الموارد البشرية",
    jobs="https://www.adwanigh.com/careers", web="https://www.adwanigh.com",
    li=NA, x=NA, mail="في صفحة التواصل بالموقع", contact="",
    found="بحثت عن مستشفيات الطائف الخاصة وأقسام IT فيها.",
    s1="https://www.adwanigh.com/departments/detail/information-technology-department",
    s2="https://www.adwanigh.com/careers", s3="",
    nxt="قدّم يدوياً واعرض عليهم بناء لوحة مؤشرات (Dashboard) تشغيلية كقيمة مضافة.")

add(p="ب", ar="جمعية الدعوة والإرشاد بالطائف", en="Dawah Society – Taif",
    sector="غير ربحي", typ="غير ربحي", size="صغير إلى متوسط",
    loc="الطائف",
    has="محتمل قوي", ev="دليل أوصاف وظيفية رسمي",
    evd="دليل الأوصاف الوظيفية المنشور على موقع الجمعية (2025) فيه بطاقة وظيفة (محلل بيانات وتقارير المدير التنفيذي).",
    date="يوليو 2025",
    roles="محلل بيانات وتقارير",
    dept="مكتب المدير التنفيذي",
    fit="عالية",
    cond="الرواتب في القطاع غير الربحي غالباً أقل، لكن الدخول فيه أسهل والخبرة تُحسب.",
    how="تقديم يدوي أو بالإيميل للجمعية",
    jobs=NA, web="https://dawahtaif.org.sa", li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن جمعيات الطائف ووظائف البيانات.",
    s1="https://dawahtaif.org.sa/wp-content/uploads/sites/596/2025/07/%D8%AF%D9%84%D9%8A%D9%84-%D8%A7%D9%84%D8%A3%D9%88%D8%B5%D8%A7%D9%81-%D8%A7%D9%84%D9%88%D8%B8%D9%8A%D9%81%D9%8A%D8%A9.pdf",
    s2="", s3="",
    nxt="اسأل إذا كانت الوظيفة مشغولة. ولو كانت شاغرة، فهي مناسبة كبداية خبرة.")

add(p="ج", ar="السلام لصناعة الطيران (مجموعة SAMI)", en="Alsalam Aerospace Industries (SAMI)",
    sector="دفاع وطيران", typ="خاص / شبه حكومي (SAMI)", size="كبير",
    loc="الطائف (صيانة الطائرات)",
    has="ممكن", ev="إعلانات تقنية في الطائف (ليست بيانات)",
    evd="إعلانات في الطائف لمسميات Fabrications Supervisor وAircraft Mechanic وDepth Support Supervisor. ما لقيت مسمى بيانات، لكن الصيانة الجوية تعتمد على بيانات الصيانة والتخطيط.",
    date="2025 – 2026",
    roles="لا يوجد مسمى بيانات صريح",
    dept="تخطيط الصيانة / Maintenance Data (متوقع)",
    fit="منخفضة إلى متوسطة",
    cond="",
    how="منصات التوظيف (Sabbar وغيرها) + موقع SAMI",
    jobs="https://sabbar.com/en/jobs/companies/alsalam-aerospace-industries", web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن شركات الطيران العاملة مع BAE في الطائف.",
    s1="https://sabbar.com/en/jobs/companies/alsalam-aerospace-industries",
    s2="https://jobsaisle.com/job/depth-support-supervisor-alsalam-aerospace-industries-at-ta-if-1", s3="",
    nxt="ابحث في LinkedIn عن موظفي Alsalam Aerospace في الطائف بمسمى Planning أو Data.")

add(p="ج", ar="شركة تجمع مطارات الثاني (مطار الطائف الدولي)", en="Cluster2 Airports – Taif International Airport",
    sector="طيران مدني", typ="شبه حكومي (مطارات القابضة)", size="كبير",
    loc="مطار الطائف",
    has="ممكن", ev="استنتاج",
    evd="الشركة تشغّل مطار الطائف الحالي منذ 2022. ومشروع المطار الجديد مطروح بنموذج شراكة بين القطاعين العام والخاص لمدة 30 سنة، يعني مشغّلاً جديداً واحتياجاً مستقبلياً. أقسام البيانات على الأغلب في المقر الرئيسي وليست في الطائف.",
    date="2022 – 2026",
    roles="غير ظاهر",
    dept="العمليات / الأداء التشغيلي",
    fit="منخفضة حالياً",
    cond="",
    how="موقع الشركة + LinkedIn",
    jobs=NA, web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن مشغّل مطار الطائف الجديد.",
    s1="https://saudipedia.com/مطار-الطائف-الدولي",
    s2="https://www.okaz.com.sa/local/na/2254543", s3="",
    nxt="تابع أخبار ترسية المطار الجديد، لأن المشغّل الجديد سيوظّف في الطائف.")

add(p="ج", ar="غرفة الطائف (الغرفة التجارية)", en="Taif Chamber of Commerce",
    sector="خدمات أعمال", typ="غير ربحي / خدمي", size="متوسط",
    loc="الطائف",
    has="ممكن", ev="نشاط دراسات وبيانات اقتصادية",
    evd="وقّعت الغرفة في 1444هـ شراكة مع معهد الملك عبدالله للبحوث بجامعة الملك سعود لدراسة الفرص الاستثمارية في الطائف. والأهم من ذلك أن الغرفة مصدر لدليل منشآت الطائف الأعضاء، وهذا يساعدك تكبّر القائمة.",
    date="1444هـ",
    roles="دراسات وأبحاث (متوقع)",
    dept="الدراسات والمعلومات",
    fit="متوسطة",
    cond="",
    how="زيارة الغرفة + طلب دليل المنشآت الكبيرة الأعضاء",
    jobs=NA, web=NA, li="https://sa.linkedin.com/company/taifcci", x=NA, mail=NA, contact="",
    found="بحثت عن الجهات اللي مقرها الطائف.",
    s1="https://saudipedia.com/%D8%BA%D8%B1%D9%81%D8%A9-%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81",
    s2="https://sa.linkedin.com/company/taifcci", s3="",
    nxt="زر الغرفة واطلب قائمة المنشآت الكبيرة والمتوسطة الأعضاء. هذي أقوى أداة عندك لتوسيع البحث.")

add(p="ج", ar="جمعية العون الخيرية بالطائف", en="Al-Aoun Charity – Taif",
    sector="غير ربحي", typ="غير ربحي", size="صغير إلى متوسط",
    loc="الطائف",
    has="ممكن", ev="إعلان سابق",
    evd="إعلان سابق لوظائف إدارية بالطائف يشترط القدرة على التعامل مع البيانات وتحليلها واستخدام البرامج الإحصائية.",
    date="قديم (التاريخ غير محدد)",
    roles="إداري مع تحليل بيانات",
    dept="", fit="متوسطة", cond="",
    how="الموقع الرسمي للجمعية + تقديم يدوي",
    jobs=NA, web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن جمعيات الطائف ووظائف البيانات.",
    s1="https://tawten.co/archives/31442", s2="", s3="",
    nxt="تواصل معهم لمعرفة الاحتياج الحالي.")

add(p="ج", ar="جمعية البر الخيرية بالطائف", en="Al-Birr Charity – Taif",
    sector="غير ربحي", typ="غير ربحي", size="متوسط",
    loc="الطائف",
    has="ممكن", ev="إعلانات إدارية سابقة",
    evd="أعلنت سابقاً عن وظائف إدارية للجنسين في الطائف. والجمعيات الكبيرة تحتاج تقارير أثر ومؤشرات لحوكمة المركز الوطني لتنمية القطاع غير الربحي.",
    date="قديم",
    roles="إداري (غير صريح)",
    dept="التخطيط والجودة / التقارير", fit="متوسطة", cond="",
    how="تقديم يدوي",
    jobs=NA, web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن جمعيات الطائف ووظائف البيانات.",
    s1="https://www.wdeftksa.com/sa/jobs/8979", s2="", s3="",
    nxt="اعرض عليهم دور (أخصائي تقارير ومؤشرات أداء).")

add(p="ج", ar="شركة الطائف للاستثمار والسياحة", en="Taif Investment & Tourism Co.",
    sector="سياحة واستثمار", typ="خاص (مساهمة مقفلة، رأس مال 110 مليون)", size="متوسط",
    loc="شارع شبرا – الطائف",
    has="ممكن", ev="لا يوجد دليل بيانات",
    evd="من الشركات القليلة اللي مقرها الطائف فعلاً، وأعلنت سابقاً عن وظائف إدارية (سكرتير براتب 5000–8000). ما لقيت دليلاً على وظائف بيانات.",
    date="غير محدد",
    roles="لا يوجد", dept="الإدارة / المالية", fit="منخفضة", cond="",
    how="تقديم يدوي في المقر",
    jobs=NA, web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن الشركات اللي مقرها الرئيسي في الطائف.",
    s1="https://www.findsaudi.com/ar/CatView/579/2416/%D8%B4%D8%B1%D9%83%D8%A9_%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81_%D9%84%D9%84%D8%A5%D8%B3%D8%AA%D8%AB%D9%85%D8%A7%D8%B1_%D9%88%D8%A7%D9%84%D8%B3%D9%8A%D8%A7%D8%AD%D8%A9.aspx",
    s2="", s3="",
    nxt="أولوية منخفضة. قدّم لهم فقط إذا كنت تمر بالمقر أصلاً.")

add(p="ج", ar="شركة إدارة للتنمية والتطوير", en="Eadarah for Development",
    sector="تطوير عقاري", typ="خاص", size="غير معروف",
    loc="شارع وادي وج – مجمع فالي سنتر – الطائف",
    has="ممكن", ev="لا يوجد دليل بيانات",
    evd="شركة مقرها الطائف، وما لقيت دليلاً على وظائف بيانات.",
    date="", roles="", dept="", fit="منخفضة", cond="",
    how="تقديم يدوي",
    jobs=NA, web="https://www.eadarah.sa", li=NA, x=NA, mail="في صفحة التواصل: https://www.eadarah.sa/Contact_Us/Index/1", contact="",
    found="بحثت عن الشركات اللي مقرها الرئيسي في الطائف.",
    s1="https://www.eadarah.sa/Contact_Us/Index/1", s2="", s3="",
    nxt="أولوية منخفضة.")

add(p="ج", ar="الشركة السعودية للكهرباء — الطائف", en="Saudi Electricity Company – Taif",
    sector="مرافق / كهرباء", typ="شبه حكومي", size="كبير جداً",
    loc="الطائف (مكاتب ومرافق)",
    has="ممكن", ev="استنتاج فقط",
    evd="ما لقيت إعلان بيانات في الطائف. وظائف البيانات في الشركة غالباً في المقرات الإقليمية مثل جدة والرياض، لكن عندها فروع تشغيل في الطائف.",
    date="", roles="", dept="العمليات / العدادات / خدمة العملاء", fit="منخفضة", cond="",
    how="بوابة الشركة + LinkedIn",
    jobs=NA, web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت عن المرافق الكبيرة في الطائف.",
    s1="", s2="", s3="",
    nxt="ابحث في LinkedIn عن (Saudi Electricity Taif) بمسمى Data أو Analyst قبل ما تصرف وقتاً عليها.")

add(p="ج", ar="مكتب استشارات (غير مسمّى) — وظيفة إحصائي", en="(Unnamed consultancy – Statistician, Taif)",
    sector="استشارات", typ="خاص", size="غير معروف",
    loc="الطائف",
    has="مؤكد (إعلان)", ev="إعلان سابق",
    evd="ظهر في نتائج Sabbar إعلان (Statistician) في الطائف براتب 4000 ريال شهرياً وخبرة 0–1 سنة. اسم الجهة ما ظهر في نتيجة البحث.",
    date="غير محدد (2026 تقريباً)",
    roles="إحصائي", dept="", fit="عالية (خبرة 0–1)",
    cond="الراتب منخفض نسبياً (4000).",
    how="ابحث في Sabbar عن (إحصائي) في الطائف",
    jobs="https://sabbar.com/ar/jobs/c-الطائف", web=NA, li=NA, x=NA, mail=NA, contact="",
    found="بحثت في Sabbar عن (Taif data analyst).",
    s1="https://sabbar.com/en/jobs/c-taif", s2="", s3="",
    nxt="افتح Sabbar وابحث عن (Statistician Taif) لمعرفة اسم الجهة.")

# ---------------------------------------------------------------- excluded / weak
EXCL = [
    ("صيدليات الدواء — الطائف", "Al-Dawaa Medical Services", "إعلانات إدخال بيانات فقط في الطائف (خبرة 2–5 سنوات). مستبعدة لأنك استثنيت إدخال البيانات، وفريق التحليل عندهم في المقر الرئيسي.", "https://sabbar.com/en/jobs/c-taif-r-data-entry/id-de226445-08a5-4b3b-9378-8872de1c95b0"),
    ("جمعية الطائف الخيرية لذوي الاحتياجات الخاصة", "Taif Charity for Special Needs", "الإعلان الوحيد كان إدخال بيانات في جدة ضمن شبكة قادرون (انتهى في 5 مايو 2023).", "https://www.taifsn.org/jobs/"),
    ("هامات (ذا بارك مول الطائف)", "Hamat", "مشغّل مولات مقره خارج الطائف. فرع الطائف تشغيلي فقط.", "https://aawsat.com/home/article/3796446"),
    ("المدينة الصناعية بالطائف (مدن)", "MODON – Taif Industrial City", "مدينة جديدة (2019) وتبعد حوالي 58 كم، ومصانعها صغيرة في الغالب (مثل 20 مصنعاً جاهزاً بمساحة 450م²). احتمال وجود فرق بيانات فيها ضعيف حالياً، لكن راجع قائمة المصانع لاحقاً.", "https://modon.gov.sa/ar/Partners/Factories/Pages/default.aspx"),
    ("قاعدة الملك فهد الجوية / القوات الجوية", "RSAF – King Fahd Air Base", "التوظيف العسكري له مسار مختلف تماماً. الفرص المدنية فيها تمر عن طريق المقاولين مثل BAE (انظر الورقة الرئيسية).", ""),
]

# ---------------------------------------------------------------- search channels
def g(q):
    return "https://www.google.com/search?q=" + quote_plus(q)

CHANNELS = [
    ("LinkedIn — أشخاص (X-ray)", "موظفو بيانات في الطائف (إنجليزي)", g('site:linkedin.com/in ("Taif" OR "Ta\'if") ("Data Analyst" OR "Business Intelligence" OR "Data Engineer" OR "Data Specialist" OR "Data Management")'), "افتح كل ملف وسجّل جهة العمل الحالية. إذا كانت جهة في الطائف ومو موجودة في القائمة فأضفها."),
    ("LinkedIn — أشخاص (X-ray)", "موظفو بيانات في الطائف (عربي)", g('site:linkedin.com/in "الطائف" ("محلل بيانات" OR "أخصائي بيانات" OR "ذكاء الأعمال" OR "إدارة البيانات")'), ""),
    ("LinkedIn — أشخاص (من حسابك)", "بحث مباشر داخل LinkedIn", "https://www.linkedin.com/search/results/people/?keywords=data%20analyst%20Taif", "فلتر Locations = Taif. وجرّب كلمات: BI، Power BI، Reporting، Informatics، إحصاء."),
    ("LinkedIn — وظائف", "وظائف data في الطائف", "https://www.linkedin.com/jobs/search/?keywords=data&location=Taif%2C%20Makkah%20Region%2C%20Saudi%20Arabia", "فعّل Job Alert."),
    ("إعلانات قديمة (Google)", "إعلانات بيانات قديمة في الطائف قبل 2025", g('"الطائف" ("محلل بيانات" OR "تحليل بيانات" OR "ذكاء الأعمال" OR "إحصائي") before:2025-01-01'), "المعامل before: يطلع لك الإعلانات القديمة. غيّر التاريخ."),
    ("إعلانات قديمة (Google)", "إعلانات قديمة بالإنجليزي", g('"Taif" ("data analyst" OR "BI" OR "data management" OR "statistician") job before:2025-06-01'), ""),
    ("Sabbar", "كل وظائف الطائف", "https://sabbar.com/ar/jobs/c-الطائف", "أكثر موقع ظهر فيه إعلانات الطائف أثناء بحثي، وفيه إعلانات قديمة كثيرة."),
    ("Sabbar (Google)", "إعلانات بيانات الطائف في Sabbar", g('site:sabbar.com "c-taif" (data OR analyst OR statistician OR informatics OR BI)'), ""),
    ("X (تويتر)", "وسم وظائف الطائف", "https://x.com/search?q=%23%D9%88%D8%B8%D8%A7%D8%A6%D9%81_%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81%20%D8%A8%D9%8A%D8%A7%D9%86%D8%A7%D8%AA&f=live", "كثير من شركات الطائف تعلن في X فقط."),
    ("جدارات", "المنصة الوطنية الموحدة للتوظيف (حكومي)", "https://jadarat.sa", "ابحث عن (بيانات)، (إحصاء)، (نظم معلومات) واختر الطائف."),
    ("BAE Systems", "بوابة BAE الرسمية", "https://jobsearch.baesystems.com", "Location: Taif، Keyword: data."),
    ("الخدمات الطبية للقوات المسلحة", "بوابة مستشفيات القوات المسلحة", "https://talents.msd.med.sa", ""),
    ("تجمع الطائف الصحي", "بوابة التجمع", "https://altaif-careers.health.sa", "التسجيل لمدة 15 يوماً فقط عند فتح الجولة."),
    ("الحرس الوطني الصحي", "بوابة الشؤون الصحية بالحرس", "https://careers.mngha.med.sa", "فلتر City = Taif."),
    ("مدن", "قائمة مصانع المدن الصناعية", "https://modon.gov.sa/ar/Partners/Factories/Pages/default.aspx", "اختر الطائف."),
]

METHOD = [
    ("الهدف", "حصر الجهات في الطائف اللي عندها قسم أو موظفين بيانات (تحليل، هندسة، BI، إدارة بيانات، إحصاء، معلوماتية صحية)، مع استبعاد إدخال البيانات."),
    ("تاريخ البحث", TODAY),
    ("الطريقة", "1) بدأت من القطاعات اللي تحتاج بطبيعتها فرق بيانات (صحة، حكومي، دفاع، مرافق، تعليم، غير ربحي)، ولم أعتمد على وجود كلمة (الطائف) في اسم الجهة. 2) بحثت عن إعلانات قديمة وحالية. 3) بحثت عن موظفين حاليين وسابقين بمسميات بيانات عن طريق LinkedIn X-ray. 4) رجعت للوائح والهياكل الرسمية. 5) استبعدت المنشآت الصغيرة ووظائف إدخال البيانات."),
    ("درجات الثقة", "مؤكد = دليل مباشر (إعلان بيانات في الطائف أو موظف بمسمى بيانات في الجهة). محتمل قوي = لائحة أو هيكل رسمي يفرض وجود وحدة بيانات، أو إعلان قريب. محتمل = جهة كبيرة والجهة الأم عندها أقسام بيانات. ممكن = جهة متوسطة أو كبيرة بدون دليل مباشر."),
    ("الأولوية", "أ = ابدأ بها هذا الأسبوع. ب = خلال الشهر. ج = متى ما سمح الوقت. والأولوية مبنية على قوة الدليل وملاءمة الجهة لحديث تخرج وإمكانية التقديم من الطائف."),
    ("قيد 1 — الوصول", "بيئة العمل منعت فتح Sabbar وLinkedIn مباشرة، فاعتمدت على نتائج محرك البحث. لهذا بعض التفاصيل (مثل اسم الشركة في بعض الإعلانات) ناقصة ومكتوب أمامها (غير مسمّى)."),
    ("قيد 2 — البريد والهاتف", "ما كتبت أي إيميل أو رقم ما تحققت منه. الخانات المكتوب فيها (لم أتحقق منه) تحتاج تراجع الموقع الرسمي. تخمين إيميل غلط يضيّع عليك فرصة."),
    ("قيد 3 — مصادر تجميع", "ZoomInfo وRocketReach وApollo بياناتها تقديرية وقد تكون قديمة، فاعتبرها مؤشرات لا حقائق."),
    ("قيد 4 — الشمول", "ما فيه مصدر واحد يجمع كل جهات الطائف، والقائمة ما تقدر تكون كاملة 100%. ورقة (قنوات البحث) فيها روابط جاهزة تكمّل فيها البحث بنفس المنهجية، وأقوى مصدر لتوسيع القائمة هو دليل أعضاء غرفة الطائف."),
    ("ملاحظة واقعية", "وظائف البيانات في الطائف تتركز في: الصحة (التجمع الصحي والمستشفيات العسكرية والحرس)، والجامعة، والجهات الحكومية (الأمانة وهيئة التطوير والتعليم)، والدفاع (BAE). والقطاع الخاص المحلي فيه احتياج قليل. لهذا التقديم اليدوي يفيد أكثر مع المستشفيات الخاصة والجمعيات والمقاولين، أما الحكومي فيمر غالباً عن طريق جدارات أو التواصل الشخصي."),
]

# ---------------------------------------------------------------- styles
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
BODY_FONT = Font(name="Arial", size=10)
LINK_FONT = Font(name="Arial", size=10, color="0563C1", underline="single")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="right", readingOrder=2)
CENTER = Alignment(wrap_text=True, vertical="top", horizontal="center", readingOrder=2)

wb = Workbook()
ws = wb.active
ws.title = "الجهات"
ws.sheet_view.rightToLeft = True

keys = ["_id", "p", "ar", "en", "sector", "typ", "size", "loc", "has", "ev", "evd", "date", "roles", "dept",
        "fit", "cond", "how", "jobs", "web", "li", "x", "mail", "contact", "found", "s1", "s2", "s3", "nxt",
        "_status", "_date", "_notes"]
assert len(keys) == len(COLS)

for ci, (h, w) in enumerate(COLS, 1):
    c = ws.cell(row=1, column=ci, value=h)
    c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CENTER, BORDER
    ws.column_dimensions[get_column_letter(ci)].width = w
ws.row_dimensions[1].height = 34

link_keys = {"jobs", "web", "li", "x", "s1", "s2", "s3"}
for ri, row in enumerate(R, 2):
    row["_id"] = ri - 1
    row["_status"] = "لم أقدّم"
    for ci, k in enumerate(keys, 1):
        v = row.get(k, "")
        c = ws.cell(row=ri, column=ci, value=v)
        c.font, c.alignment, c.border = BODY_FONT, WRAP, BORDER
        if k in ("_id", "p", "has", "fit", "_status", "_date"):
            c.alignment = CENTER
        if k in link_keys and isinstance(v, str) and v.startswith("http"):
            c.hyperlink = v
            c.font = LINK_FONT
    ws.row_dimensions[ri].height = 150

last = len(R) + 1
ref = f"A1:{get_column_letter(len(COLS))}{last}"
tab = Table(displayName="TaifDataEmployers", ref=ref)
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
ws.add_table(tab)
ws.freeze_panes = "D2"

# data validation for status
status_col = get_column_letter(keys.index("_status") + 1)
dv = DataValidation(type="list", formula1='"لم أقدّم,قدّمت إلكترونياً,قدّمت يدوياً,تواصلت مع شخص,مقابلة,عرض,مرفوض,مؤجل"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"{status_col}2:{status_col}500")
date_col = get_column_letter(keys.index("_date") + 1)
for r in range(2, last + 1):
    ws[f"{date_col}{r}"].number_format = "yyyy-mm-dd"

# conditional colors
has_col = get_column_letter(keys.index("has") + 1)
pri_col = get_column_letter(keys.index("p") + 1)
rng_has = f"{has_col}2:{has_col}500"
for txt, color in [("مؤكد", "C6EFCE"), ("محتمل قوي", "DDEBF7"), ("محتمل", "FFF2CC"), ("ممكن", "F2F2F2")]:
    ws.conditional_formatting.add(rng_has, FormulaRule(formula=[f'LEFT({has_col}2,{len(txt)})="{txt}"'], fill=PatternFill("solid", fgColor=color), stopIfTrue=True))
rng_p = f"{pri_col}2:{pri_col}500"
for txt, color in [("أ", "63BE7B"), ("ب", "FFEB84"), ("ج", "D9D9D9")]:
    ws.conditional_formatting.add(rng_p, FormulaRule(formula=[f'{pri_col}2="{txt}"'], fill=PatternFill("solid", fgColor=color)))

def simple_sheet(title, headers, widths, rows, link_cols=()):
    s = wb.create_sheet(title)
    s.sheet_view.rightToLeft = True
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = s.cell(row=1, column=ci, value=h)
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, CENTER, BORDER
        s.column_dimensions[get_column_letter(ci)].width = w
    for ri, r in enumerate(rows, 2):
        for ci, v in enumerate(r, 1):
            c = s.cell(row=ri, column=ci, value=v)
            c.font, c.alignment, c.border = BODY_FONT, WRAP, BORDER
            if ci in link_cols and isinstance(v, str) and v.startswith("http"):
                c.hyperlink = v
                c.font = LINK_FONT
    s.freeze_panes = "A2"
    if rows:
        s.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows)+1}"
    return s

simple_sheet("قنوات البحث", ["القناة", "الغرض", "الرابط الجاهز", "طريقة الاستخدام"], [26, 40, 70, 60], CHANNELS, link_cols=(3,))
simple_sheet("مستبعدة أو ضعيفة", ["الجهة", "الاسم الإنجليزي", "سبب الاستبعاد", "المصدر"], [34, 30, 80, 60], EXCL, link_cols=(4,))
simple_sheet("المنهجية والقيود", ["البند", "التفاصيل"], [22, 130], METHOD)

# summary sheet with formulas
sm = wb.create_sheet("ملخص", 0)
sm.sheet_view.rightToLeft = True
sm["A1"] = "ملخص جهات البيانات في الطائف"
sm["A1"].font = Font(bold=True, size=14, name="Arial")
sm["A2"] = f"تاريخ البحث: {TODAY} — عدد الجهات في الورقة الرئيسية = {len(R)}"
sm["A2"].font = BODY_FONT
rows = [("حسب درجة الثقة", None)]
sm["A4"], sm["B4"] = "درجة الثقة", "العدد"
sm["D4"], sm["E4"] = "الأولوية", "العدد"
sm["G4"], sm["H4"] = "حالة التقديم", "العدد"
for c in ("A4", "B4", "D4", "E4", "G4", "H4"):
    sm[c].fill, sm[c].font, sm[c].alignment = HDR_FILL, HDR_FONT, CENTER
for i, t in enumerate(["مؤكد*", "محتمل قوي", "محتمل", "ممكن"], 5):
    sm[f"A{i}"] = t.replace("*", "")
    sm[f"B{i}"] = f"=COUNTIF('الجهات'!{has_col}:{has_col},\"{t}\")"
for i, t in enumerate(["أ", "ب", "ج"], 5):
    sm[f"D{i}"] = t
    sm[f"E{i}"] = f"=COUNTIF('الجهات'!{pri_col}:{pri_col},\"{t}\")"
for i, t in enumerate(["لم أقدّم", "قدّمت إلكترونياً", "قدّمت يدوياً", "تواصلت مع شخص", "مقابلة", "عرض", "مرفوض", "مؤجل"], 5):
    sm[f"G{i}"] = t
    sm[f"H{i}"] = f"=COUNTIF('الجهات'!{status_col}:{status_col},\"{t}\")"
for col, w in zip("ABCDEFGH", [18, 10, 3, 12, 10, 3, 20, 10]):
    sm.column_dimensions[col].width = w
sm["A15"] = "طريقة الاستخدام: افتح ورقة (الجهات)، واستخدم أسهم الفلترة في رأس الأعمدة (الأولوية، الثقة، القطاع، الملاءمة)، وحدّث عمود (حالة التقديم) من القائمة المنسدلة."
sm["A15"].alignment = WRAP
sm.merge_cells("A15:H17")

wb.save(OUT)
print("saved", OUT, len(R), "rows")
