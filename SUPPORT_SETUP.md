# إعداد نظام البلاغات

هذا الملف يشرح طريقة ربط زر الدعم داخل البرنامج بـ Google Apps Script بدون وضع رابط الاستقبال داخل الكود المفتوح أو README.

## 1. إنشاء Google Sheet أو Apps Script

1. افتح حساب الدعم الخاص بك.
2. أنشئ Google Sheet جديد باسم مناسب مثل `Marriage Manager Support`، ثم من القائمة اختر `Extensions` ثم `Apps Script`.
3. يمكنك أيضا فتح Apps Script مباشرة من `script.google.com` بدون Google Sheet. في هذه الحالة سينشئ السكربت ملف Sheet تلقائيا عند أول بلاغ.
4. انسخ محتوى [support_google_apps_script.gs](support_google_apps_script.gs) داخل محرر Apps Script.
5. غيّر البريد التالي إلى بريد الدعم الحقيقي:

```js
const SUPPORT_RECIPIENT_EMAIL = "your-support-email@example.com";
```

## 2. نشر Web App

1. من Apps Script اختر `Deploy` ثم `New deployment`.
2. اختر النوع `Web app`.
3. اضبط الإعدادات كالتالي:
   - `Execute as`: Me
   - `Who has access`: Anyone
4. اضغط `Deploy`.
5. وافق على الصلاحيات المطلوبة.
6. انسخ رابط `Web app URL`.

## 3. ربط البرنامج بشكل خاص

لا تضع رابط `Web app URL` داخل الكود أو داخل README أو داخل إعدادات عامة. الرابط يعتبر إعدادا خاصا بالمالك فقط.

### عند البناء المحلي

قبل بناء النسخة الجاهزة، عرّف الرابط كمتغير بيئة:

```powershell
$env:MARRIAGE_MANAGER_SUPPORT_URL = "ضع رابط Web App هنا"
.\build_portable.ps1
.\build_installer.ps1
```

سيقوم سكربت البناء بإنشاء ملف `support_endpoint.txt` محليا وإدخاله داخل النسخة المبنية. هذا الملف موجود في `.gitignore` ولا يجب رفعه إلى GitHub.

### عند البناء عبر GitHub Actions

أضف Secret في GitHub باسم:

```text
MARRIAGE_MANAGER_SUPPORT_URL
```

ثم ضع فيه رابط `Web app URL`. عند إنشاء Release من tag، سيتم حقن الرابط داخل النسخة الجاهزة فقط، بدون ظهوره في الكود المفتوح.

## ملاحظات أمنية

- البرنامج يقبل فقط روابط `https://script.google.com/macros/s/.../exec` كوجهة دعم.
- سكربت Google Apps Script يضع حدودا لطول الحقول ويتحقق من نوع الرسالة والبريد الإلكتروني.
- السكربت يحمي Google Sheets من Formula Injection عند تسجيل البلاغات.
- إخفاء الرابط عن GitHub والواجهة يحميه من الظهور العرضي، لكنه لا يجعله سرا مطلقا داخل تطبيق Desktop مبني. للحماية الأقوى استخدم Backend خاصا يتحقق من الطلبات قبل تمريرها.

## أنواع الرسائل

- `مشكلة تقنية`: خطأ، توقف، حساب غير صحيح، أو شيء لا يعمل.
- `طلب ميزة`: إضافة وظيفة جديدة غير موجودة حاليا.
- `اقتراح تحسين`: تحسين على وظيفة أو شاشة موجودة.
- `ملاحظة عامة`: رأي أو تعليق لا يدخل بوضوح تحت الأنواع السابقة.
