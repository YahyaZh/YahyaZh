# ---- Manual-application sheet: established Taif companies (advertise or not)
PITCH = {
 'تجزئة وتجارة': ('تحليل المبيعات والمخزون، ولوحات مؤشرات للفروع', 'محلل بيانات مبيعات / أخصائي تقارير'),
 'صحة': ('تقارير مؤشرات الجودة (CBAHI)، والمطالبات التأمينية، وإشغال الأسرّة', 'أخصائي بيانات/تقارير صحية'),
 'تعليم': ('تقارير التسجيل والأداء الأكاديمي، وتنظيم بيانات الطلاب', 'أخصائي بيانات وتقارير'),
 'ضيافة وسياحة': ('تحليل الإشغال والإيرادات (Occupancy/RevPAR)، وتقارير الحجوزات', 'محلل إيرادات / تقارير'),
 'مقاولات/هندسة': ('ضبط المشاريع والتكاليف، وتقارير التقدم', 'أخصائي ضبط مشاريع / تقارير'),
 'صناعة/أغذية': ('تقارير الإنتاج والجودة والتوزيع', 'محلل بيانات تشغيلية'),
 'مطاعم ومقاهي': ('تحليل مبيعات الفروع وبيانات نقاط البيع (POS)', 'محلل بيانات مبيعات'),
 'تشغيل وصيانة': ('مؤشرات مستوى الخدمة في العقود (SLA)، وتقارير أوامر العمل', 'أخصائي تقارير تشغيلية'),
 'غير ربحي': ('تقارير الأثر ومؤشرات الحوكمة للمركز الوطني لتنمية القطاع غير الربحي', 'أخصائي قياس وتقارير'),
 'تقنية': ('حلول بيانات للعملاء، وتحليل وهندسة البيانات', 'محلل/مهندس بيانات'),
 'مرافق وطاقة': ('تقارير التشغيل والعدادات والشكاوى', 'محلل بيانات تشغيلية'),
 'نقل ولوجستيات': ('تحليل الشحنات والأسطول والتوصيل', 'محلل بيانات لوجستية'),
 'مالية': ('تقارير مالية وتحليل المحافظ', 'محلل بيانات مالية'),
}
DEFAULT_PITCH = ('تنظيم البيانات وبناء تقارير ولوحات مؤشرات (Excel/Power BI)', 'أخصائي بيانات وتقارير')

DIRECTORY_COS = [  # found via directories/news, few or no job ads
 ("مؤسسة نجم الطائف (مجموعة سونار)", "مقاولات/هندسة", "525 موظفاً و302 مشروعاً (حسب موقعها)", "https://najem.sa/about/", "دليل/موقع الشركة"),
 ("شركة شبرا الطائف للخدمات اللوجستية", "نقل ولوجستيات", "غير معروف", "https://www.dnb.com/business-directory/company-information.general_freight_trucking.sa.makkah_al_mukkaramah.at_taif.html", "دليل Dun & Bradstreet"),
 ("شركة الجزيرة العربية المحدودة للتجارة والصناعة والمقاولات", "مقاولات/هندسة", "غير معروف", "https://www.dnb.com/business-directory/company-information.general_freight_trucking.sa.makkah_al_mukkaramah.at_taif.html", "دليل Dun & Bradstreet"),
 ("الماجد التجارية", "تجزئة وتجارة", "201–500 (LinkedIn)", "https://www.linkedin.com/company/almajed524046208", "LinkedIn"),
 ("مصنع الطائف الوطني لمنتجات الألبان", "صناعة/أغذية", "51–200 (LinkedIn) وإيرادات حوالي 18 مليون دولار (ZoomInfo)", "https://sa.linkedin.com/company/al-taif-national-dairy-factory", "LinkedIn + ZoomInfo"),
 ("مخابز وحلويات مذاق الملكة", "صناعة/أغذية", "متوسط (مصنع بتقييم +A)", "https://www.linkedin.com/company/queenstaste", "LinkedIn + هيئة الغذاء والدواء"),
 ("شركة القاضي للزراعة والدواجن", "صناعة/أغذية", "متوسط", "https://www.sfda.gov.sa/en/node/16962", "هيئة الغذاء والدواء"),
 ("عنب الطائف (منتجات ورق العنب)", "صناعة/أغذية", "غير معروف", "https://www.findsaudi.com/ar/CatView/267/12024.aspx", "Find Saudi + إعلان محاسب"),
 ("خيرات الطائف", "صناعة/أغذية", "صغير إلى متوسط", "https://www.findsaudi.com/ar/CatView/243/23328/%D8%AE%D9%8A%D8%B1%D8%A7%D8%AA_%D8%A7%D9%84%D8%B7%D8%A7%D8%A6%D9%81.aspx", "Find Saudi"),
 ("براق الطائف للأسمدة السائلة", "صناعة/أغذية", "غير معروف", "https://www.findsaudi.com/def/Search.aspx?q=&cityId=1&pagerId=530", "Find Saudi"),
 ("نخب الطائف للخرسانة الجاهزة", "صناعة/أغذية", "غير معروف", "https://readymix.sa/ar/saudi/makkah/taif", "Readymix.sa"),
 ("Ibrand (اي براند) للأدوية", "صحة", "51–200 (LinkedIn)", "https://sa.linkedin.com/company/ibrandsa", "LinkedIn"),
 ("شركة المساند لتقنية المعلومات", "تقنية", "11–50 (LinkedIn)", "https://sa.linkedin.com/company/al-musand", "LinkedIn"),
 ("بنزول (Benzol) لتجارة المنتجات البترولية", "تجزئة وتجارة", "غير معروف (تأسست 2018)", "https://sa.linkedin.com/company/benzol", "LinkedIn"),
 ("BINGO (أغذية ومشروبات)", "مطاعم ومقاهي", "11–50 (LinkedIn)", "https://www.linkedin.com/company/bracket-sa", "LinkedIn"),
 ("شركة تازة (مطاعم، أكثر من 22 فرعاً)", "مطاعم ومقاهي", "متوسط", "https://www.linkedin.com/company/taza-company", "LinkedIn"),
 ("المدرسة الباكستانية العالمية بالطائف", "تعليم", "201–500 (LinkedIn)", "https://www.linkedin.com/company/pakistan-international-school-taif", "LinkedIn"),
 ("فندق إنتركونتيننتال الطائف", "ضيافة وسياحة", "201–500 (LinkedIn)، 179 غرفة", "https://sn.linkedin.com/company/intercontinental-taif", "LinkedIn"),
 ("فندق لو ميريديان الهدا", "ضيافة وسياحة", "111 غرفة", "https://www.eyeofriyadh.com/ar/directory/details/6023_le-m-ridien-al-hada-hotel", "دليل"),
 ("فندق وأجنحة رمادا الهدا", "ضيافة وسياحة", "متوسط", "https://www.eyeofriyadh.com/ar/directory/details/6022_ramada-al-hada-hotel-and-suites", "دليل"),
 ("هامات (ذا بارك مول الطائف)", "تجزئة وتجارة", "كبير (المقر خارج الطائف)", "https://aawsat.com/home/article/3796446", "أخبار"),
 ("شركة إدارة للتنمية والتطوير", "مقاولات/هندسة", "غير معروف", "https://www.eadarah.sa/Contact_Us/Index/1", "موقع الشركة"),
 ("مجموعة السفياني لأنظمة الأمن والسلامة", "مقاولات/هندسة", "متوسط (تأسست 2012)", "https://x.com/alsufyanigroup", "X + Find Saudi"),
 ("مجموعة ماتركس", "تقنية", "صغير إلى متوسط", "https://matrixs.com.sa/", "موقع الشركة"),
 ("G7 Software House", "تقنية", "صغير", "https://www.linkedin.com/company/g7-soft", "LinkedIn"),
 ("امتياز العربية للاستشارات", "تقنية", "غير معروف (فيها مدير بنية تحتية IT)", "https://www.linkedin.com/in/al-mqdashi-abdulmughni-0380791b0/", "حساب موظف"),
 ("مجموعة بن شافي التعليمية", "تعليم", "متوسط إلى كبير", "https://www.linkedin.com/in/ahmed-awad-16529526/", "حساب موظف"),
 ("مطاعم الحاشي البري", "مطاعم ومقاهي", "متوسط (عندها مدير مالي)", "https://www.linkedin.com/in/mohamed-heshmat-1949a2195/", "حساب موظف"),
 ("مستشفى الأمين العام", "صحة", "حوالي 100 سرير", "https://www.zavis.ai/ar/sa/directory/taif/hospitals/al-ameen-hospital", "دليل + إعلان إحصائي"),
 ("مستشفى العدواني العام", "صحة", "80–100 سرير", "https://www.adwanigh.com/careers", "موقع المستشفى"),
]

PRIVATE_EXCL = re.compile(r'جامعة|أمانة|امانة|وزارة|هيئة|الهيئة|المؤسسة العامة|الخدمات الطبية|القوات المسلحة|الحرس|الداخلية|تجمع|ديوان|بنك الرياض|البنك|مصرف|جدارات|برنامج|أكاديمية|معهد الإدارة', re.I)
NATIONAL = re.compile(r'المراعي|بنده|النهدي|الدواء|أرامكو|سابك|STC|موبايلي|زين|جرير|عبداللطيف جميل|بيبسيكو|ماكدونالدز|كودو|البيك|هرفي|ايكيا|إيكيا|extra|إكسترا|سدافكو|نادك|الراجحي|Bechtel|IHG|Chalhoub|Arla|Danube|الدانوب|أباريل|Abdul Latif|SEVEN|مشاريع الترفيه|Al Rushaid|BAE|بي أيه إي|سراكو|SRACO|تحكم|العرض المتقن|بن لادن|الخطوط|طيران', re.I)

man_rows = []
seenk = set()
for (n, sec, size, link, src) in DIRECTORY_COS:
    p = PITCH.get(sec, DEFAULT_PITCH)
    man_rows.append([n, sec, size, 'لا أو قليل', '', src, link, p[0], p[1]])
    seenk.add(re.sub(r'[^\w]', '', n)[:8])
for o in sorted(emps, key=lambda o: -o['n']):
    n = o['name']
    if o['n'] < 2 or PRIVATE_EXCL.search(n) or NATIONAL.search(n) or re.fullmatch(r'(شركة|مؤسسة|مجمع|مركز|مطاعم|مدارس|مجموعة)( (طبي|طبية|أسنان|طبي خاص|محاماة|حجامة|تغذية|كبرى|خاصة|عالمية|للمقاولات|مقاولات|تجارية|أطفال|للاطفال|أسنان طبي))*', n.strip()) or len(n.split())<2 or 'العامة للري' in n:
        continue
    k = re.sub(r'[^\w]', '', n)[:8]
    if k in seenk:
        continue
    seenk.add(k)
    sec = sector(n)
    p = PITCH.get(sec, DEFAULT_PITCH)
    ex = (o['ex1'] + o['ex2'] + o['ex'])[:1]
    man_rows.append([n, sec, f"مؤشر: {o['n']} إعلاناً في الطائف", f"نعم ({o['n']})", f"{o['first']} → {o['last']}", "، ".join(o['srcs']), ex[0][2] if ex else '', p[0], p[1]])
for i, r in enumerate(man_rows, 1):
    r.insert(0, i)
    r += ['لم أقدّم', '', '']
mh = ["م", "الجهة", "القطاع (تقديري)", "الحجم / مؤشر الحجم", "هل تعلن؟ (عدد الإعلانات)", "فترة الإعلانات", "مصدر المعلومة", "رابط", "ليش ممكن تحتاج بيانات؟", "الدور اللي تعرضه", "حالة التقديم", "تاريخ الزيارة/التقديم", "ملاحظاتي"]
mw = [6, 40, 16, 30, 16, 24, 24, 40, 50, 28, 14, 14, 30]
sheet("شركات الطائف للتقديم اليدوي", mh, mw, man_rows, link_cols={8}, center_cols={1, 5, 11, 12}, table_name="ManualApply", row_h=45, status_col=11, freeze="C2")
N_MANUAL = len(man_rows)
