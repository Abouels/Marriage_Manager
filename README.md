<div dir="rtl" align="right">

# مدير مصاريف الزواج

تطبيق سطح مكتب عربي لإدارة مصاريف تجهيز الشقة والزواج. البرنامج يحفظ البيانات محليا على جهازك، ويدعم البنود، الفواتير، التوريدات، التسديدات، السلف، التقارير، والنسخ الاحتياطي.

</div>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white">
  <img alt="Tkinter" src="https://img.shields.io/badge/UI-Tkinter-0F766E?style=for-the-badge">
  <img alt="SQLite" src="https://img.shields.io/badge/Database-SQLite-003B57?style=for-the-badge&logo=sqlite&logoColor=white">
  <img alt="Windows" src="https://img.shields.io/badge/Platform-Windows-2563EB?style=for-the-badge&logo=windows&logoColor=white">
</p>

<div dir="rtl" align="right">

## التحميل والتشغيل

لو أنت مستخدم عادي، لا تحتاج تثبيت Python ولا تحميل أي مكتبات برمجة. حمل النسخة الجاهزة من صفحة:

[GitHub Releases](https://github.com/Abouels/Marriage_Manager/releases)

كل إصدار رسمي يوفر اختيارين:

| الملف | مناسب لمن؟ | طريقة التشغيل |
| --- | --- | --- |
| `MarriageExpensesManager_Setup_*.exe` | لو عايز تثبيت عادي على ويندوز | حمل الملف، افتحه، واتبع خطوات التثبيت |
| `MarriageExpensesManager_Portable_*.zip` | لو عايز نسخة تفتح مباشرة بدون تثبيت | حمل الملف، فك الضغط، ثم افتح `MarriageExpensesManager.exe` |

### تثبيت نسخة Setup

1. افتح صفحة [Releases](https://github.com/Abouels/Marriage_Manager/releases).
2. حمل ملف `MarriageExpensesManager_Setup_*.exe` من أحدث إصدار.
3. افتح ملف التسطيب.
4. بعد انتهاء التثبيت افتح البرنامج من Start Menu أو من الاختصار على سطح المكتب.

بيانات نسخة التسطيب تحفظ داخل مجلد المستخدم في ويندوز:

```text
%LOCALAPPDATA%\MarriageExpensesManager\app_data
```

### تشغيل نسخة Portable

1. افتح صفحة [Releases](https://github.com/Abouels/Marriage_Manager/releases).
2. حمل ملف `MarriageExpensesManager_Portable_*.zip` من أحدث إصدار.
3. فك الضغط في أي مكان مناسب، مثل سطح المكتب أو فلاشة USB.
4. افتح `MarriageExpensesManager.exe` من داخل المجلد بعد فك الضغط.

بيانات النسخة المحمولة تحفظ بجانب ملف التشغيل داخل نفس المجلد:

```text
MarriageExpensesManager/app_data
```

> ملاحظة: لو ظهرت رسالة Windows SmartScreen لأن البرنامج غير موقع رقميا بعد، اختر `More info` ثم `Run anyway` إذا كنت حملت الملف من صفحة Releases الرسمية.

## المميزات

| القسم | الوصف |
| --- | --- |
| المصاريف | تسجيل بنود الفرش، التشطيب، حفل الزواج، التوريدات، والمصاريف الأخرى |
| الحسابات | دعم نظام الطرف الواحد، نظام 50/50، ونظام تحديد المسؤول لكل بند |
| المرفقات | إرفاق فواتير وملفات لكل بند مع دعم السحب والإفلات عند توفر المكتبة |
| التوريدات | تسجيل المبالغ المستلمة وطريقة الدفع والتاريخ والملاحظات |
| التسديدات | متابعة المدفوعات لجهات الدفع المختلفة |
| السلف | تسجيل السلف بعملات متعددة ومتابعة المسدد والمتبقي |
| التقارير | تصدير Excel وتقارير PDF |
| النسخ الاحتياطي | تصدير واستيراد نسخة احتياطية من بيانات البرنامج |

## لقطات من الواجهة

<p align="center">
  <img src="assets/screenshots/dashboard.png" alt="لوحة إدارة مصاريف الزواج" width="900">
</p>

<table>
  <tr>
    <td align="center">
      <strong>عرض التوريدات</strong><br>
      <img src="assets/screenshots/receipts.png" alt="عرض التوريدات" width="420">
    </td>
    <td align="center">
      <strong>إدارة السلف</strong><br>
      <img src="assets/screenshots/loans.png" alt="إدارة السلف" width="420">
    </td>
  </tr>
</table>

## للمطورين

هذا القسم فقط لمن يريد تشغيل السورس أو تعديل الكود. المستخدم العادي يفضل تحميل النسخة الجاهزة من Releases.

ثبت المتطلبات:

```powershell
py -m pip install -r requirements.txt
```

شغل البرنامج أثناء التطوير:

```powershell
py app.py
```

أو شغله بدون نافذة Console:

```powershell
py app.pyw
```

`app.py` هو ملف التطبيق الأساسي وفيه الكود والمنطق والواجهة.  
`app.pyw` مجرد launcher صغير يستدعي `main()` من `app.py` لتشغيل البرنامج على ويندوز بدون نافذة Console.

## إصدار نسخة رسمية

النسخ الجاهزة للمستخدمين يتم بناؤها تلقائيا عبر GitHub Actions عند إنشاء tag مثل:

```powershell
git tag v1.0.0
git push origin v1.0.0
```

بعد دفع الـ tag، سيقوم GitHub Actions ببناء:

```text
MarriageExpensesManager_Setup_1.0.0.exe
MarriageExpensesManager_Portable_v1.0.0.zip
```

ثم يرفق الملفات في صفحة GitHub Release الخاصة بالإصدار.

## هيكل المشروع

```text
apartment_manager/
|-- app.py                         # التطبيق الأساسي
|-- app.pyw                        # Launcher صامت لويندوز
|-- requirements.txt               # متطلبات التطوير
|-- build_portable.ps1             # بناء نسخة Portable للمطورين
|-- build_installer.ps1            # بناء Installer للمطورين
|-- packaging/
|   |-- MarriageExpensesManager.spec
|   `-- MarriageExpensesManager.iss
|-- .github/workflows/
|   `-- release-windows.yml        # بناء ملفات Releases تلقائيا
|-- assets/
|   |-- icons/
|   `-- screenshots/
`-- app_data/                      # بيانات محلية غير مرفوعة على GitHub
```

## ملاحظات مهمة

- لا ترفع `app_data` إلى GitHub لأنه يحتوي على قواعد بيانات وملفات مستخدمين.
- عدل الواجهة والمنطق داخل `app.py` فقط، واترك `app.pyw` كـ launcher.
- قبل عمل إصدار رسمي، اختبر:

```powershell
py -m py_compile app.py app.pyw
```

## صاحب المشروع

Designed & Developed by Eng. I. Abouelsaad  
i.abouelsaad9@gmail.com

</div>
