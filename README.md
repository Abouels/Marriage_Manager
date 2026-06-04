<div dir="rtl" align="right">

# مدير مصاريف الزواج

تطبيق سطح مكتب عربي لإدارة مصاريف تجهيز الشقة وحفل الزفاف، مع نظام حسابات مرن، قاعدة بيانات محلية، مرفقات فواتير، وتصدير تقارير منظمة.

</div>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white">
  <img alt="Tkinter" src="https://img.shields.io/badge/UI-Tkinter-0F766E?style=for-the-badge">
  <img alt="SQLite" src="https://img.shields.io/badge/Database-SQLite-003B57?style=for-the-badge&logo=sqlite&logoColor=white">
  <img alt="Windows" src="https://img.shields.io/badge/Platform-Windows-2563EB?style=for-the-badge&logo=windows&logoColor=white">
</p>

<div dir="rtl" align="right">

## نظرة سريعة

`Marriage Manager` يساعدك تسجل كل بند مالي في تجهيز الشقة والزفاف، وتعرف مين دفع، مين عليه، المتبقي كام، وإيه تفاصيل السلف والتوريدات. كل البيانات محفوظة محليا على جهازك داخل SQLite، بدون خادم أو تسجيل دخول.

## لقطات من الواجهة

### لوحة الإدارة الرئيسية

<p align="center">
  <img src="assets/screenshots/dashboard.png" alt="لوحة إدارة مصاريف الزواج" width="900">
</p>

### القوائم المالية

<table>
  <tr>
    <td align="center">
      <strong>التوريدات المسجلة</strong><br>
      <img src="assets/screenshots/receipts.png" alt="قائمة التوريدات المسجلة" width="420">
    </td>
    <td align="center">
      <strong>إدارة السلف</strong><br>
      <img src="assets/screenshots/loans.png" alt="نافذة إدارة السلف" width="420">
    </td>
  </tr>
</table>

## ماذا يفعل البرنامج؟

| القسم | الوصف |
| --- | --- |
| المصاريف | تسجيل الفرش، التشطيب، حفل الزفاف، والمصاريف الأخرى |
| الحسابات | دعم نظام الطرف الواحد، نظام 50/50، ونظام البنود على الطرفين |
| التوريدات | تسجيل المبالغ المستلمة وطريقة الدفع والتاريخ والملاحظات |
| التسديدات | متابعة المدفوعات لجهات الدفع مثل والد العروسة أو مهندس التشطيب |
| السلف | تسجيل السلف بعملات متعددة ومتابعة المسدد والمتبقي |
| المرفقات | حفظ فواتير ومرفقات لكل بند أو سداد |
| التقارير | تصدير Excel وتقارير PDF عند توفر مكتبات PDF |

## أنظمة الحسابات

| النظام | الاستخدام |
| --- | --- |
| نظام الطرف الواحد | كل البنود الأساسية على الطرف الأول، مع سياسة خاصة لحفل الزفاف |
| نظام 50/50 | البنود المشتركة تتحسب مناصفة بين الطرفين |
| نظام بنود على الطرفين | كل بند يتم تحديد الطرف المسؤول عنه مباشرة |

## التشغيل على Windows

### 1. تثبيت المتطلبات

```powershell
py -m pip install -r requirements.txt
```

### 2. تشغيل البرنامج

للتشغيل العادي:

```powershell
.\launch_windows.bat
```

للتشغيل مع نافذة أخطاء أثناء التطوير:

```powershell
.\launch_debug.bat
```

أو تشغيل مباشر:

```powershell
py app.py
```

## هيكل المشروع

```text
Marriage_Manager/
|-- app.py                 # نسخة تشغيل مع Console للتطوير وتتبع الأخطاء
|-- app.pyw                # نسخة تشغيل صامتة مناسبة للانشر Windows
|-- requirements.txt       # مكتبات Python المطلوبة
|-- launch_windows.bat     # تشغيل سريع
|-- launch_debug.bat       # تشغيل Debug
|-- launch_windows.vbs     # تشغيل صامت عبر Windows Script Host
|-- assets/
|   `-- icons/             # أيقونات بطاقات الملخص
`-- app_data/              # بيانات محلية غير مرفوعة على GitHub
```

## البيانات والخصوصية

البرنامج ينشئ ويحفظ البيانات محليا داخل:

```text
app_data/apartment_costs.db
```

هذا المجلد متجاهل من Git عن قصد، لأنه يحتوي على قاعدة البيانات، الفواتير، الإيصالات، وملفات التصدير الخاصة بالمستخدم.

## الملفات المحلية التي ينشئها البرنامج

| المسار | المحتوى |
| --- | --- |
| `app_data/apartment_costs.db` | قاعدة بيانات SQLite |
| `app_data/invoices` | مرفقات وفواتير البنود |
| `app_data/loan_receipts` | إيصالات سداد السلف |
| `app_data/exports` | ملفات Excel و PDF المصدرة |
| `app_data/startup_error.log` | سجل أخطاء بدء التشغيل عند الحاجة |

## المتطلبات

| المكتبة | الاستخدام |
| --- | --- |
| `openpyxl` | تصدير ملفات Excel |
| `Pillow` | عرض ومعالجة الصور والمرفقات |
| `tkinterdnd2` | دعم السحب والإفلات عند توفره |
| `reportlab` | إنشاء تقارير PDF |
| `arabic-reshaper` و `python-bidi` | تحسين عرض العربية داخل PDF |

## ملاحظات للمطورين

- التطبيق يعتمد على Tkinter و SQLite فقط، لذلك تشغيله بسيط ولا يحتاج Backend.
- `app.py` و `app.pyw` يجب أن يظلا متزامنين لأن اللانشر الصامت يستخدم `app.pyw`.
- لا ترفع `app_data` إلى GitHub حتى لا يتم نشر بيانات المستخدم.
- عند تعديل منطق الحسابات، راجع دوال مشاركة الأطراف والتقارير معا حتى تظل الشاشة والتصدير متطابقين.

## حالة المشروع

النسخة الحالية جاهزة للاستخدام المحلي، وقاعدة البيانات تبدأ فارغة في التثبيت الجديد. إعداد أول تشغيل يظهر فقط عند عدم وجود إعدادات محفوظة.

</div>
