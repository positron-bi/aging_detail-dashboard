# گزارش حساب‌های دریافتنی و وصول — Python

فایل `FactFinnance1.xlsx` فکت منبع است و فقط روی سیستم محلی نگه داشته می‌شود. خروجی Excel مطابق ساختار فایل الگوی «گزارش حساب‌های دریافتنی و وصول ماه ۶ پنج شرکت» تولید می‌شود.

## اجرای داشبورد

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

مسیر پیش‌فرض فکت محلی:

```text
C:\Users\a.farshchian\Desktop\AR_Aging_Report\FactFinnance1.xlsx
```

در صورت تغییر مسیر، متغیر محیطی `FACT_FINANCE_PATH` را تنظیم کنید یا فایل را در داشبورد انتخاب کنید.

## تولید مستقیم Excel

```bash
python report_generator.py --input FactFinnance1.xlsx --output AR_Collections_Report.xlsx --year 1405 --month 6
```

## تولید محلی و آپلود خودکار خروجی

```bash
python generate_and_upload.py --year 1405 --month 6
```

این فرمان ابتدا گزارش را داخل پوشه `reports` تولید می‌کند و سپس فقط همان خروجی تولیدشده را commit و روی شاخه `main` پوش می‌کند. فایل فکت به‌وسیله `.gitignore` از Git خارج نگه داشته می‌شود.

برای تولید بدون آپلود:

```bash
python generate_and_upload.py --year 1405 --month 6 --no-push
```

خروجی شامل شیت‌های خلاصه، ریز فروش، ریز وصول، ریز وصول اسناد، استثناها، کنترل گزارش و منبع گردش کل است. مبالغ گزارش به میلیون تومان نمایش داده می‌شوند ولی محاسبات با مبلغ دقیق ریالی انجام می‌شوند.

فکت ستون مستقل «شماره سند/عطف» ندارد؛ برنامه `FinID` را به‌عنوان ردیف قابل ردیابی حفظ می‌کند، شماره ساختگی نمی‌سازد و اتصال‌های غیرقطعی را در شیت استثناها گزارش می‌دهد.
