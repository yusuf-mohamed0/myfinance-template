---
name: MyFinance System
description: نظام الميزانية الشخصية - تحليل SMS بنكي، الخطة الشهرية، الأسعار الأسبوعية، حالة الموقع والنشر. استخدم عند أي سؤال عن المال أو الميزانية أو المصروفات أو الأسعار أو تشغيل سكربتات المشروع.
---

# نظام MyFinance - دليل التشغيل

## خريطة المجلدات

| المسار | المحتوى |
|---|---|
| `01_Data/sms_raw*.txt` | SMS الخام (الأحدث mtime هو المعتمد) |
| `02_Reports/` | `summary_<year>.json`، `transactions_<year>.csv`، `sms_report.txt`، `market_watch.json`، `charts/` |
| `03_System/config.json` | إعدادات المستخدم (حر) |
| `03_System/source_manifest.json` | بصمات موقّعة لملفات النظام (مقفول) |
| `04_Source/` | سكربتات النظام (مقفولة - لا تُعدَّل أبدًا) |
| `05_Docs/` | README + الوثائق |
| `06_Web/` | الموقع المولد + مستودع GitHub Pages |

## خط الإنتاج

```bash
python3 04_Source/analyze_sms.py              # SMS -> تقارير
python3 04_Source/plan_month.py <income> <YYYY-MM>   # خطة الشهر
python3 04_Source/build_web.py                # بناء الموقع
python3 04_Source/publish_web.py              # نشر (اسأل قبلها)
python3 04_Source/verify_system.py            # الفحص النهائي - لازم أخضر
python3 04_Source/daily_status.py             # حالة اليوم / المتبقي الأسبوعي
```

## قواعد لا تخالفها

1. كل ملفات `04_Source/` و`init_project.py` و`source_manifest.json` و`AGENTS.md` و`.opencode/` و`setup.*` **مقفولة ببصمة موقّعة** — أي تعديل = `SYSTEM FILES MODIFIED`. التطوير يحتاج إذن صاحب النظام.
2. بعد أي تعديل مسموح: شغّل `verify_system.py` ولا تعلن الإتمام قبل خضرائه.
3. أسعار السوق تتحدث **كل سبت** في `market_watch.json` — ولا تضع رمز `→` (U+2192) فيه أبدًا.
4. اليوم يبدأ من **السبت** (حساب الأسابيع والـcaps).
5. لا تحذف ملفات SMS ولا تعدّل أرقام التقارير يدويًا.
6. الردود باللهجة المصرية العربية من غير إيموجي، والأسعار والنسب بالأرقام الدقيقة.
7. اسأل قبل `git push` / `gh *` / أي نشر.
