<div dir="rtl" align="right">

# مدير مصاريف الزواج

تطبيق سطح مكتب عربي لإدارة مصاريف تجهيز الشقة والزواج، مبني بـ Python و Tkinter و SQLite. البرنامج يحفظ البيانات محليا على الجهاز، ويدعم البنود، الفواتير، التوريدات، التسديدات، السلف، التقارير، والنسخ الاحتياطي.

</div>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white">
  <img alt="Tkinter" src="https://img.shields.io/badge/UI-Tkinter-0F766E?style=for-the-badge">
  <img alt="SQLite" src="https://img.shields.io/badge/Database-SQLite-003B57?style=for-the-badge&logo=sqlite&logoColor=white">
  <img alt="Windows" src="https://img.shields.io/badge/Platform-Windows-2563EB?style=for-the-badge&logo=windows&logoColor=white">
</p>

<div dir="rtl" align="right">

## المميزات

| القسم | الوصف |
| --- | --- |
| المصاريف | تسجيل بنود الفرش، التشطيب، حفل الزواج، التوريدات، والمصاريف الأخرى |
| الحسابات | دعم نظام الطرف الواحد، نظام 50/50، ونظام تحديد المسؤول لكل بند |
| المرفقات | إرفاق فواتير وملفات لكل بند مع دعم السحب والإفلات عند توفر المكتبة |
| التوريدات | تسجيل المبالغ المستلمة وطريقة الدفع والتاريخ والملاحظات |
| التسديدات | متابعة المدفوعات لجهات الدفع المختلفة |
| السلف | تسجيل السلف بعملات متعددة ومتابعة المسدد والمتبقي |
| التقارير | تصدير Excel وتقارير PDF عند توفر المكتبات المطلوبة |
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

## تشغيل المشروع من السورس

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

## الفرق بين app.py و app.pyw

`app.py` هو ملف التطبيق الأساسي وفيه كل الكود والمنطق والواجهة. استخدمه أثناء التطوير لأنه يظهر الأخطاء في الـ Console.

`app.pyw` مجرد launcher صغير يستدعي `main()` من `app.py`. وجوده مهم لويندوز ولـ PyInstaller لأنه يشغل البرنامج بدون نافذة Console سوداء. لا تكرر الكود داخله.

## بناء نسخة Portable

يشترط وجود Python على جهاز المطور فقط. السكربت يثبت PyInstaller إذا لم يكن موجودا، ثم يبني نسخة جاهزة داخل `dist`.

```powershell
.\build_portable.ps1
```

الناتج:

```text
dist/MarriageExpensesManager/MarriageExpensesManager.exe
```

النسخة المحمولة تحتوي على `portable.flag` ومجلد `app_data` بجانب ملف التشغيل، لذلك تحفظ بيانات المستخدم داخل نفس مجلد البرنامج.

## بناء نسخة Installer

ثبت Inno Setup 6 أولا، ثم شغل:

```powershell
.\build_installer.ps1
```

الناتج:

```text
dist/installer/
```

نسخة التسطيب تحفظ بيانات المستخدم في:

```text
%LOCALAPPDATA%\MarriageExpensesManager\app_data
```

## مسارات البيانات

أثناء التطوير من السورس:

```text
app_data/
```

في النسخة المحمولة:

```text
MarriageExpensesManager/app_data/
```

في النسخة المثبتة:

```text
%LOCALAPPDATA%\MarriageExpensesManager\app_data
```

يمكن تغيير مسار البيانات بتعيين المتغير:

```powershell
$env:MARRIAGE_MANAGER_DATA_DIR="D:\MyData\MarriageExpenses"
```

## هيكل المشروع

```text
apartment_manager/
|-- app.py                         # التطبيق الأساسي
|-- app.pyw                        # Launcher صامت لويندوز
|-- requirements.txt               # المتطلبات
|-- build_portable.ps1             # بناء نسخة Portable
|-- build_installer.ps1            # بناء Installer عبر Inno Setup
|-- packaging/
|   |-- MarriageExpensesManager.spec
|   `-- MarriageExpensesManager.iss
|-- assets/
|   |-- icons/
|   `-- screenshots/
`-- app_data/                      # بيانات محلية غير مرفوعة على GitHub
```

## ملاحظات للمطورين

- لا ترفع `app_data` إلى GitHub لأنه يحتوي على قواعد بيانات وملفات مستخدمين.
- عدل الواجهة والمنطق داخل `app.py` فقط، واترك `app.pyw` كـ launcher.
- قبل النشر شغل:

```powershell
py -m py_compile app.py app.pyw
.\build_portable.ps1
```

## صاحب المشروع

Designed & Developed by Eng. I. Abouelsaad  
i.abouelsaad9@gmail.com

</div>
