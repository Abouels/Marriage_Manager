# إعداد نظام البلاغات

## 1. إنشاء Google Sheet أو Apps Script

1. افتح حساب الدعم الخاص بك.
2. أنشئ Google Sheet جديد باسم مناسب مثل `Marriage Manager Support`، ثم من القائمة اختر `Extensions` ثم `Apps Script`.
3. لو فتحت Apps Script مباشرة من `script.google.com` بدون Google Sheet، سيعمل السكربت أيضًا وينشئ الشيت تلقائيًا عند أول بلاغ.
4. انسخ محتوى ملف `support_google_apps_script.gs` داخل محرر Apps Script.
5. لو عايز البلاغات توصل لإيميل مختلف، غير قيمة:

```js
const SUPPORT_RECIPIENT_EMAIL = "your-support-email@example.com";
```

## 2. نشر Web App

1. من Apps Script اختر `Deploy` ثم `New deployment`.
2. اختر النوع `Web app`.
3. اضبط:
   - `Execute as`: Me
   - `Who has access`: Anyone
4. اضغط `Deploy`.
5. وافق على الصلاحيات المطلوبة.
6. انسخ رابط `Web app URL`.

## 3. ربط البرنامج بشكل خاص

لا تضع رابط `Web app URL` داخل الكود أو داخل README أو في إعدادات البرنامج. الرابط يعتبر إعدادًا خاصًا بالمالك فقط.

### عند البناء المحلي

قبل بناء النسخة الجاهزة، عرّف الرابط كمتغير بيئة:

```powershell
$env:MARRIAGE_MANAGER_SUPPORT_URL = "ضع رابط Web App هنا"
.\build_portable.ps1
.\build_installer.ps1
```

سيقوم سكربت البناء بإنشاء ملف `support_endpoint.txt` محليًا وإدخاله داخل النسخة المبنية. هذا الملف موجود في `.gitignore` ولا يجب رفعه إلى GitHub.

### عند البناء عبر GitHub Actions

أضف Secret في GitHub باسم:

```text
MARRIAGE_MANAGER_SUPPORT_URL
```

ثم ضع فيه رابط `Web app URL`. عند إنشاء Release من tag، سيتم حقن الرابط داخل النسخة الجاهزة فقط، بدون ظهوره في الكود المفتوح.

> ملاحظة أمنية: إخفاء الرابط عن الواجهة والكود المفتوح يحميه من المستخدم العادي ومن GitHub، لكن أي رابط مدمج داخل تطبيق Desktop يمكن استخراجه نظريًا بتحليل ملفات النسخة المبنية. لو احتجت حماية أقوى جدًا، استخدم Backend خاص يتحقق من صلاحيات الطلبات قبل تمرير البلاغات.

## أنواع الرسائل

- `مشكلة تقنية`: خطأ، توقف، حساب غير صحيح، أو شيء لا يعمل.
- `طلب ميزة`: إضافة وظيفة جديدة غير موجودة حاليًا.
- `اقتراح تحسين`: تحسين على وظيفة أو شاشة موجودة.
- `ملاحظة عامة`: رأي أو تعليق لا يدخل بوضوح تحت الأنواع السابقة.
