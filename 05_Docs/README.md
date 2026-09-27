# MyFinance — نظام الإدارة المالية الشخصي (قالب جاهز)

نظام محلي بالكامل: تحليل SMS بنكي، ميزانية شهرية بالجنيه، حالة يومية، موقع ويب محمي بباسورد.
كل البيانات تعيش على جهازك، والحساب على GitHub بتاعك أنت.

---

## المتطلبات

| المتطلب | ويندوز | ماك (macOS) |
|---|---|---|
| بايثون 3.10+ | python.org | `brew install python3` |
| مكتبات | `pip install openpyxl python-docx plotly` | `pip3 install openpyxl python-docx plotly` |
| GitHub CLI | winget install GitHub.cli | `brew install gh` |
| تسجيل دخول GitHub | `gh auth login` | `gh auth login` (بحسابك أنت) |

---

## التشغيل من الصفر (مرة واحدة)

```bash
# 1) انسخ المجلد كامل لأي مكان (مثلا ~/myfinance)
# 2) شغّل سكربت الإعداد — يسألك الاسم والباسورد والدخل ويجهز كل حاجة:
python3 init_project.py

# أو بالأعلام من غير أسئلة:
python3 init_project.py --yes --name "اسمك" --passcode "باسوردك" \
    --income 30000 --city "القاهرة" --no-repo
```

سكربت الإعداد بيعمل بالترتيب:

1. يكتب `03_System/config.json` (الاسم، الباسورد، الدخل، العملة، المدينة، السنة، GitHub).
2. يولّد ملف SMS تجريبي ديناميكي (`01_Data/sms_raw_sample.txt`) عشان خط الإنتاج يشتغل فورًا — **امسحه وحط ملفك الحقيقي بنفس الاسم أو أي اسم يبدأ بـ `sms_raw`**.
3. يبني: إكسل → تحليل → خطة الشهر → الموقع → فحص (`verify_system`).
4. لو اسم المستخدم/الريبو متاحين على GitHub: يعمل repo خاص عادي Pages مش بيشتغل عليه → ينشئه، يفعّل Pages، ويدفع.

بعد النجاح: افتح `06_Web/index.html` محليًا، وحط الباسورد.

---

## الأوامر اليومية

```bash
cd 04_Source
python3 build_excel.py              # تحديث ملف الإكسل
python3 analyze_sms.py              # تحليل ملفات 01_Data/sms_raw*.txt (الأحدث mtime)
python3 plan_month.py               # خطة الشهر الحالي من ملخص التحليل
python3 daily_status.py             # حالة اليوم: المصروف vs الميزانية المتبقية
python3 build_web.py                # توليد 06_Web/index.html
python3 publish_web.py              # بناء + دفع + انتظار نشر Pages
python3 verify_system.py            # فحص النظام كامل (٦٠+ فحص)
```

الترتيب المعتاد (مرة في الشهر / بعد أي تحديث):

```bash
python3 build_excel.py && python3 analyze_sms.py && python3 plan_month.py \
  && python3 build_web.py && python3 verify_system.py && python3 publish_web.py
```

> المكتبات مثبتة؟ شغّل الأمر ده لو ظهر `ModuleNotFoundError`:
> `pip3 install openpyxl python-docx plotly`

---

## إضافة بيانات SMS جديدة

1. انسخ رسائل SMS البنك لملف نصي: `01_Data/sms_raw.txt` (أو `sms_raw_2026-10.txt`).
2. الملف الأحدث بآخر تعديل هو اللي بيتحلل — ملفات قديمة بتتجاهل تلقائيًا.
3. الصيغة المتوقعة (بنك مصر / NBE):

```
تم خصم 250.00 EGP عند SHELL cairo يوم 15/9/2026 ... المتاح 4750.00
تم إضافة تحويل داخلي مبلغ 10000.00 من مكافأة رقم مرجعي 123 يوم 20/9/2026
تم رد مبلغ 50.00 EGP ... يوم 21/9/2026
ناسف لعدم إتمام المعاملة بمبلغ 90.00 EGP
```

- رسائل OTP و«ناسف لعدم إتمام» ورسائل النظام بتتشال تلقائيًا.
- الردود (refunds) بتتطرح من المصروف.
- **بنك تاني؟** التعديل في مكانين بس داخل `04_Source/analyze_sms.py`:
  - تعبيرات `R_*` (سطور ~35–55) — نمط الرسالة.
  - دالة `categorize()` (سطور ~67–96) — تصنّف الاسم لمظلة.
  ده **تعديل بيانات مش تعديل نظام**، ومينفعش يتراجع من غير إذن صاحب النظام.

---

## تعديل الميزانية الشهرية

`04_Source/plan_month.py` → قاموس `ENVELOPES` (نسب مئوية لازم مجموعها = 100%، فيه assert بيرفض غير كده).
كمان `market_watch.json` بيربط أسعار السوق بالموازنة — راجع قسم «نبض السوق» في الموقع.

---

## أمان النظام (مهم)

- `03_System/source_manifest.json` فيه بصمة SHA-256 لكل ملف في `04_Source/` + `init_project.py`.
- `verify_system.py` بيفحص البصمات مع كل تشغيل — أي تعديل في ملفات السيستيم يظهر:
  `SYSTEM FILES MODIFIED` والفحص بيفشل.
- **ممنوع تعديل ملفات `04_Source/` أو `init_project.py` من غير إذن صاحب النظام.**
  البيانات (config، SMS، تقارير، إكسل) تُعدَّل بحرية — القفل على الكود فقط.
- لإعادة توليد البصمات بعد تعديل مُصرّح عليه:

```bash
python3 -c "import hashlib,json,pathlib;b=pathlib.Path('..');f=sorted(list((b/'04_Source').glob('*.py')))+[b/'init_project.py'];(b/'03_System'/'source_manifest.json').write_text(json.dumps({str(p.relative_to(b)).replace(chr(92),'/'):hashlib.sha256(p.read_bytes().replace(b'\r\n',b'\n')).hexdigest() for p in f},indent=2))"
```

---

## النشر على GitHub Pages

1. `gh auth login` بحسابك.
2. `python3 publish_web.py` — يبني ويدفع وينتظر اكتمال Pages، ويطبع الرابط.
3. **تنبيه خصوصية:** GitHub Pages المجاني لازم الريبو يكون **عام**. يعني الكود والملفات المنشورة (إكسل، تقارير، موقع) متاحة لأي حد عنده الرابط — الموقع نفسه محمي بباسورد على المتصفح، لكن الباكدج اللِيڤل مش مخفي.
   - لو مش عايز ده: خلي الريبو private واستخدم `index.html` محلي بس، أو ادفع لـ Pages Pro.
4. الباسورد في `config.json` → بيتحوّل لبصمة SHA-256 داخل `index.html` (الباسورد نفسه مش بيتطبع في الكود أبدًا).

---

## محتويات المجلد

```
MyFinance_Template/
├─ init_project.py          إعداد أول مرة (بيولد الـ manifest وقت التشغيل لو مفيش)
├─ 01_Data/                 ملفات SMS الخام (أنت بتحطها هنا)
├─ 02_Reports/              ناتج التحليل + market_watch.json (أسعار أسبوعية)
│   ├─ sms_analysis.json
│   ├─ sms_report.txt
│   ├─ market_watch.json
│   └─ charts/              رسوم plotly
├─ 03_System/
│   ├─ config.json          إعداداتك (اسم/باسورد/دخل/مدينة/ريبو)
│   ├─ source_manifest.json بصمات ملفات النظام
│   └─ My_Financial_System.xlsx   مصنوع من build_excel
├─ 04_Source/               كود النظام (مقفول بالبصمات)
│   ├─ mfconfig.py          تحميل الإعدادات المشتركة
│   ├─ build_excel.py       بناء ملف الإكسل
│   ├─ analyze_sms.py       تحليل SMS + JSON/تقرير/رسوم
│   ├─ plan_month.py        خطة الشهر من الملخص
│   ├─ daily_status.py      حالة اليوم
│   ├─ build_web.py         توليد الموقع
│   ├─ publish_web.py       دفع Pages
│   ├─ build_word.py        دليل PDF/Word
│   └─ verify_system.py     الفحص الشامل
├─ 05_Docs/                 الدليل ده + دليل الإدارة (docx)
└─ 06_Web/                  index.html المولّد (هنا git repo لو فعّلت GitHub)
```

---

## حل المشاكل

| المشكلة | الحل |
|---|---|
| `ModuleNotFoundError: openpyxl` | `pip3 install openpyxl python-docx plotly` |
| `verify ... FAILED` وذكر `SYSTEM FILES MODIFIED` | حد عدّل كود النظام — راجع `git diff` في `06_Web` أو أعد النسخة النظيفة |
| `gh: not found` | ويندوز: `winget install GitHub.cli` — ماك: `brew install gh` — وبعدها أعد الطرفية |
| الموقع بيفتح فاضي | شغّل `python3 build_web.py` — أو افتح `06_Web/index.html` مباشرة |
| رابط Pages مش شغال | استنى ~دقيقة بعد `publish_web.py`، أو `gh api repos/<owner>/<repo>/pages` |
| تحليل فاضي «لا توجد معاملات» | ملف `01_Data/sms_raw*.txt` ناقص أو صيغته مش متوافقة (شوف قسم SMS فوق) |
