# Hokok Tools V2

## الهدف
تحويل الكتاب إلى معرفة منظمة ثم محتوى تعليمي، دون ربط جودة المحتوى بالواجهة الحالية.

## خط المعالجة
PDF → Extraction → Verification → Structure → Knowledge Map → Concept Registry → Importance → Lessons → Question Bank → Deduplication → JSON/Word

## التشغيل
داخل `tools-node`: `npm install` ثم `node generate-book.js`.
ضع مفاتيح Gemini في `GEMINI_API_KEYS` وضع ملفات PDF في `input/`.

## المخرج الأساسي
`output/<book>-v2.json` وهو المصدر المناسب لاحقًا لتطبيق Android/Firebase/API.

## ملاحظات
- `generate-book-v1.js` نسخة احتياطية من الأداة القديمة.
- V2 تستخدم `CONTENT_MODE=source_strict` افتراضيًا.
- عدد الأسئلة أصبح سياسة سقف تكيفية بدل عدد إلزامي لكل جزء.
