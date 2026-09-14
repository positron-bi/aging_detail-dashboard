"""Generate the receivables/collections report from FactFinnance1.xlsx.

The output follows the supplied Persian workbook template. The source has no
accounting document/reference number; FinID is retained as the auditable row ID
and uncertain links are reported instead of invented.
"""
from __future__ import annotations

import argparse
import io
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

try:
    import xlsxwriter
except ModuleNotFoundError:  # The fallback keeps the CLI usable on a clean Python install.
    xlsxwriter = None

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
MILLION_TOMAN_RIAL = 10_000_000
COMPANY_NAMES = {"10": "الماس", "20": "گسترش", "30": "پارس", "40": "رهجو", "50": "سیمرغ"}
AR_RE = re.compile(r"حساب\s*های?\s*دریافتنی.*تجاری|حسابهای\s*دریافتنی.*تجاری", re.I)
DOC_RE = re.compile(r"اسناد\s*(?:تجاری\s*)?دریافتنی", re.I)
BANK_RE = re.compile(r"بانک|موجودی ریالی نزد بانک|واسط.*بانک|کارتخوان|درگاه|وجوه در راه|POS", re.I)
CASH_RE = re.compile(r"صندوق|تنخواه|وجه نقد", re.I)
SALE_RE = re.compile(r"فروش|درآمد فروش|فاکتور|برگشت از فروش|تخفیف", re.I)
INVOICE_RE = re.compile(r"(?:فاکتور|صورتحساب)[^0-9]{0,20}(?:شماره)?\s*([0-9]+)", re.I)
CHECK_RE = re.compile(r"(?:چک|صیادی)[^0-9]{0,12}([0-9]{3,})", re.I)


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _number(value) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _date(value) -> str:
    digits = re.sub(r"\D", "", _text(value))[:8]
    return digits.zfill(8) if digits else ""


def _shown_date(value: str) -> str:
    return f"{value[:4]}/{value[4:6]}/{value[6:8]}" if len(value) == 8 else value


def _column_index(ref: str) -> int:
    match = re.match(r"[A-Z]+", ref)
    letters = match.group(0) if match else "A"
    out = 0
    for letter in letters:
        out = out * 26 + ord(letter) - 64
    return out - 1


def _cell_value(cell):
    kind = cell.get("t")
    if kind == "inlineStr":
        inline = cell.find(NS + "is")
        return "" if inline is None else "".join(inline.itertext())
    node = cell.find(NS + "v")
    if node is None or node.text is None:
        return ""
    if kind in (None, "n"):
        try:
            return float(node.text)
        except ValueError:
            pass
    return node.text


def _iter_rows(zf: zipfile.ZipFile, entry: str):
    for _event, row in ET.iterparse(zf.open(entry), events=("end",)):
        if row.tag != NS + "row":
            continue
        values = {}
        for cell in row:
            if cell.tag == NS + "c":
                values[_column_index(cell.get("r", "A1"))] = _cell_value(cell)
        yield [values.get(i, "") for i in range(max(values, default=-1) + 1)]
        row.clear()


def _find_fact_sheet(zf: zipfile.ZipFile) -> tuple[str, list[str]]:
    entries = sorted(n for n in zf.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))
    for entry in entries:
        labels = [_text(x) for x in next(_iter_rows(zf, entry), [])]
        if {"FinID", "عنوان", "Balance", "CompanyRef", "DateKey"}.issubset(labels):
            return entry, labels
    raise ValueError("شیت استاندارد دارای FinID، عنوان، Balance، CompanyRef و DateKey پیدا نشد.")


@dataclass(slots=True)
class Row:
    row_id: str
    account: str
    debit: float
    credit: float
    company: str
    date: str
    detail_code: str
    detail_type: str
    customer: str
    description: str
    account_code: str

    @property
    def customer_key(self):
        return self.company, self.detail_code, self.customer

    @property
    def match_key(self):
        description = re.sub(r"\s+", " ", self.description.replace("ي", "ی").replace("ك", "ک")).strip()
        return self.company, self.date, description


def load_source(path_or_bytes: str | Path | bytes) -> list[Row]:
    stream = io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, bytes) else path_or_bytes
    with zipfile.ZipFile(stream) as zf:
        entry, headers = _find_fact_sheet(zf)
        indexes = {name: i for i, name in enumerate(headers)}
        records: list[Row] = []
        iterator = _iter_rows(zf, entry)
        next(iterator, None)
        for values in iterator:
            def get(name):
                index = indexes.get(name, 10**9)
                return values[index] if index < len(values) else ""
            account = _text(get("عنوان"))
            if not (AR_RE.search(account) or DOC_RE.search(account) or BANK_RE.search(account) or CASH_RE.search(account) or SALE_RE.search(account)):
                continue
            company_code = _text(get("CompanyRef"))
            records.append(Row(
                row_id=_text(get("FinID")), account=account, debit=_number(get("بدهکار")), credit=_number(get("بستانکار")),
                company=COMPANY_NAMES.get(company_code, company_code), date=_date(get("DateKey")),
                detail_code=_text(get("کد تفصیل سطح 4")), detail_type=_text(get("عنوان نوع تفصیل سطح 4")),
                customer=_text(get("عنوان تفصیل سطح 4")),
                description=_text(get("شرح قلم سند حسابداری")) or _text(get("شرح سند حسابداری")), account_code=_text(get("کد"))))
    return records


def _category(account: str) -> str:
    if DOC_RE.search(account): return "اسناد دریافتنی"
    if BANK_RE.search(account): return "بانک/واسط بانکی"
    if CASH_RE.search(account): return "نقد/صندوق"
    if AR_RE.search(account): return "حساب دریافتنی"
    return "سایر"


def _best_debit(row: Row, candidates: Iterable[Row]) -> Row | None:
    options = [x for x in candidates if x.debit > 0 and x.row_id != row.row_id]
    return min(options, key=lambda x: (abs(x.debit - row.credit), 0 if _category(x.account) != "سایر" else 1)) if options else None


def analyze(records: list[Row], year: int, month: int) -> dict:
    start, end = f"{year:04d}{month:02d}01", f"{year:04d}{month:02d}31"
    current = [r for r in records if start <= r.date <= end]
    ar_all = [r for r in records if AR_RE.search(r.account)]
    docs_all = [r for r in records if DOC_RE.search(r.account)]
    by_match = defaultdict(list)
    for row in records: by_match[row.match_key].append(row)

    opening = defaultdict(float)
    for row in ar_all:
        if row.date < start: opening[row.customer_key] += row.debit - row.credit

    sales_detail, receipt_detail, doc_detail, exceptions = [], [], [], []
    sales_ids = set()
    for row in current:
        if AR_RE.search(row.account) and (INVOICE_RE.search(row.description) or SALE_RE.search(row.description)):
            invoice = INVOICE_RE.search(row.description)
            sales_detail.append([row.company,row.customer,_shown_date(row.date),invoice.group(1) if invoice else "",row.row_id,row.debit,row.credit,row.account,row.description])
            sales_ids.add(row.row_id)

    receipt_amounts = defaultdict(lambda: defaultdict(float))
    for row in current:
        if not (AR_RE.search(row.account) and row.credit > 0 and row.row_id not in sales_ids): continue
        partner = _best_debit(row, by_match[row.match_key])
        kind = _category(partner.account) if partner else "نامشخص"
        allocated = min(row.credit, partner.debit) if partner else 0
        receipt_detail.append([row.company,row.customer,_shown_date(row.date),"",row.row_id,kind,row.credit,partner.row_id if partner else "",allocated,partner.debit if partner else 0,partner.account if partner else "",partner.customer if partner else "",row.description,partner.description if partner else "","شرکت + تاریخ + شرح قلم" if partner else "بدون لینک",row.credit-allocated])
        receipt_amounts[row.customer_key][kind] += allocated if partner else row.credit
        if not partner:
            exceptions.append(["ح.د بستانکار بدون طرف مشخص",row.company,row.customer,_shown_date(row.date),"",row.row_id,row.credit,row.account,row.description])
        elif abs(row.credit-partner.debit) > max(1_000_000,row.credit*.0001):
            exceptions.append(["اختلاف مبلغ بانک و ح.د خارج از تلورانس",row.company,row.customer,_shown_date(row.date),"",row.row_id,row.credit-partner.debit,partner.account,row.description])

    historic_checks = []
    for doc in docs_all:
        if doc.debit <= 0 or doc.date >= start: continue
        check = CHECK_RE.search(doc.description)
        ar_partner = next((x for x in by_match[doc.match_key] if AR_RE.search(x.account) and x.credit > 0),None)
        historic_checks.append((check.group(1) if check else "",doc,ar_partner))

    docs_amounts = defaultdict(lambda: defaultdict(float))
    for doc in current:
        if not (DOC_RE.search(doc.account) and doc.credit > 0): continue
        bank = _best_debit(doc,by_match[doc.match_key])
        if not bank or _category(bank.account) != "بانک/واسط بانکی": continue
        check = CHECK_RE.search(doc.description); check_id = check.group(1) if check else ""
        candidates = [x for x in historic_checks if (check_id and x[0] == check_id) or abs(x[1].debit-doc.credit) <= max(1_000_000,doc.credit*.0001)]
        chosen = candidates[0] if len(candidates) == 1 else None
        customer = chosen[2].customer if chosen and chosen[2] else doc.customer
        company = chosen[2].company if chosen and chosen[2] else doc.company
        key = (company,chosen[2].detail_code if chosen and chosen[2] else doc.detail_code,customer)
        category = "قدیمی" if chosen else "نامشخص"; docs_amounts[key][category] += min(doc.credit,bank.debit)
        doc_detail.append([company,customer,_shown_date(doc.date),"",doc.row_id,doc.credit,bank.row_id,bank.debit,doc.account,bank.account,doc.description,bank.description,_shown_date(chosen[1].date) if chosen else "",chosen[1].row_id if chosen else "","شناسه چک/مبلغ و لینک تاریخی" if chosen else "منشأ اثبات نشد",category])
        if not chosen:
            exceptions.append(["چک وصول‌شده بدون مشتری قطعی",doc.company,doc.customer,_shown_date(doc.date),"",doc.row_id,doc.credit,doc.account,doc.description])
        elif len(candidates) > 1:
            exceptions.append(["چند مشتری محتمل برای یک چک",doc.company,doc.customer,_shown_date(doc.date),"",doc.row_id,doc.credit,doc.account,doc.description])

    sales_amounts = defaultdict(lambda:[0.0,0.0])
    for row in current:
        if row.row_id in sales_ids:
            sales_amounts[row.customer_key][0] += row.debit; sales_amounts[row.customer_key][1] += row.credit
    universe = set(opening)|set(sales_amounts)|set(receipt_amounts)|set(docs_amounts)
    summary = []
    for key in sorted(universe):
        company,_code,customer = key; op=opening[key]; sd,sc=sales_amounts[key]; net=sd-sc
        direct=receipt_amounts[key]["بانک/واسط بانکی"]; received_docs=receipt_amounts[key]["اسناد دریافتنی"]
        cash=receipt_amounts[key]["نقد/صندوق"]; other=receipt_amounts[key]["سایر"]+receipt_amounts[key]["نامشخص"]
        receipts=direct+received_docs; prior=min(max(op,0),receipts); current_receipt=max(0,min(max(net,0),receipts-prior)); month_balance=max(0,net-current_receipt)
        old=docs_amounts[key]["قدیمی"]; current_doc=docs_amounts[key]["جاری"]; unknown=docs_amounts[key]["نامشخص"]
        summary.append([company,customer or "بدون تفصیل",op,sd,sc,net,direct,received_docs,cash,other,receipts,prior,current_receipt,month_balance,old+current_doc+unknown,old,current_doc,unknown,"","",""])
    return {"month":month,"summary":summary,"sales":sales_detail,"receipts":receipt_detail,"docs":doc_detail,"exceptions":exceptions}


def _to_million(value): return value/MILLION_TOMAN_RIAL if isinstance(value,(int,float)) else value


def _build_workbook_openpyxl(result: dict) -> bytes:
    """Small dependency-free fallback using the openpyxl package."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook(); wb.remove(wb.active)
    month_fa = str(result["month"]).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))
    summary_name, sales_name = "خلاصه ح.د و وصول", f"ریز ح.د فروش ماه {month_fa}"
    sheets = [
        (summary_name,["شرکت","تفصیل",f"مانده ح.د اول ماه {month_fa}",f"ح.د بدهکار فروش ماه {month_fa}",f"ح.د بستانکار فروش ماه {month_fa}",f"ح.د خالص فروش ماه {month_fa}","وصول مستقیم ح.د → بانک/واسط","ح.د → اسناد دریافتنی","ح.د → نقد/صندوق","سایر تسویه ح.د","جمع وصول مشتری (بانک + اسناد)","وصول از مانده قبل (FIFO)",f"وصول از فروش ماه {month_fa}",f"مانده ح.د فروش ماه {month_fa}","وصول اسناد → بانک","از وصول اسناد: قدیمی",f"از وصول اسناد: ماه {month_fa}","از وصول اسناد: منشأ نامشخص","لینک ریز فروش","لینک ریز وصول","لینک وصول اسناد"],result["summary"]),
        (sales_name,["شرکت","تفصیل","تاریخ","شماره سند/عطف","ردیف سند","بدهکار ح.د (میلیون تومان)","بستانکار ح.د (میلیون تومان)","حساب معین","شرح"],result["sales"]),
        ("ریز وصول ح.د",["شرکت","تفصیل","تاریخ","شماره سند/عطف","ردیف ح.د","نوع طرف","بستانکار ح.د","ردیف طرف","مبلغ تخصیص‌یافته","بدهکار طرف","حساب طرف","تفصیل طرف","شرح ح.د","شرح طرف","روش لینک","اختلاف"],result["receipts"]),
        ("ریز وصول اسناد",["شرکت","تفصیل مشتری","تاریخ وصول","شماره سند/عطف","ردیف اسناد","مبلغ وصول اسناد","ردیف بانک","بدهکار بانک","حساب اسناد","حساب بانک","شرح اسناد","شرح بانک","تاریخ دریافت اولیه","سند دریافت اولیه","مبنای شناسایی مشتری","دسته"],result["docs"]),
        ("استثناها",["نوع استثنا","شرکت","تفصیل","تاریخ","شماره سند/عطف","ردیف سند","مبلغ","حساب","شرح"],result["exceptions"])]
    header_fill = PatternFill("solid", fgColor="173A5E"); header_font = Font(bold=True,color="FFFFFF")
    first_rows = {}
    for name, columns, rows in sheets:
        ws=wb.create_sheet(name[:31]); ws.sheet_view.rightToLeft=True; ws.freeze_panes="C2"; ws.auto_filter.ref=f"A1:{get_column_letter(len(columns))}{max(1,len(rows)+1)}"; first_rows[name]={}
        for col,label in enumerate(columns,1):
            c=ws.cell(1,col,label); c.fill=header_fill; c.font=header_font; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        for ridx,row in enumerate(rows,2):
            if len(row)>1:first_rows[name].setdefault((row[0],row[1]),ridx)
            for col,value in enumerate(row,1):
                monetary=(name==summary_name and 3<=col<=18) or (name==sales_name and col in (6,7)) or (name=="ریز وصول ح.د" and col in (7,9,10,16)) or (name=="ریز وصول اسناد" and col in (6,8)) or (name=="استثناها" and col==7)
                cell=ws.cell(ridx,col,_to_million(value) if monetary else value); cell.alignment=Alignment(vertical="top",wrap_text=False); cell.number_format="#,##0.000;[Red]-#,##0.000" if monetary else "General"
        for i in range(1,len(columns)+1): ws.column_dimensions[get_column_letter(i)].width=34 if i==2 else 18
    ws=wb.create_sheet("کنترل گزارش"); ws.sheet_view.rightToLeft=True
    controls=[["مانده خالص ح.د اول دوره",sum(r[2] for r in result["summary"]),"جمع مانده حساب‌های دریافتنی تجاری قبل از ماه"],["ح.د بدهکار سند فروش",sum(r[3] for r in result["summary"]),"ردیف‌های فکت دارای قرینه فروش/فاکتور"],["ح.د بستانکار سند فروش",sum(r[4] for r in result["summary"]),"برگشت/تعدیل در ردیف‌های فروش"],["ح.د خالص طرف فروش",sum(r[5] for r in result["summary"]),"بدهکار منهای بستانکار"],["وصول مستقیم بانک/واسط",sum(r[6] for r in result["summary"]),"تطبیق شرکت + تاریخ + شرح"],["ح.د → اسناد دریافتنی",sum(r[7] for r in result["summary"]),"دریافت چک از مشتری"],["ح.د → نقد/صندوق",sum(r[8] for r in result["summary"]),"وصول نقدی"],["سایر تسویه/انتقال ح.د",sum(r[9] for r in result["summary"]),"طرف غیر بانکی یا نامشخص"],["وصول اسناد → بانک",sum(r[14] for r in result["summary"]),"نقدشدن چک؛ وصول جدید مشتری نیست"],["تعداد استثناها",len(result["exceptions"]),"موارد غیرقطعی حذف نشده‌اند"]]
    headers=["کنترل","مبلغ (میلیون تومان)","توضیح"]
    for col,label in enumerate(headers,1): c=ws.cell(1,col,label); c.fill=header_fill; c.font=header_font
    for r,row in enumerate(controls,2): ws.cell(r,1,row[0]); ws.cell(r,2,_to_million(row[1]) if r<11 else row[1]); ws.cell(r,3,row[2])
    source=wb.create_sheet("منبع گردش کل"); source.sheet_view.rightToLeft=True
    for col,label in enumerate(["منبع","فایل منبع کامل","توضیح"],1): c=source.cell(1,col,label); c.fill=header_fill; c.font=header_font
    source.append(["FactFinance","FactFinnance1.xlsx","فکت منبع؛ FinID به‌عنوان ردیف قابل ردیابی حفظ شده است"]); source.append(["محدودیت","شماره سند/عطف در فکت موجود نیست","شماره ساختگی تولید نشده و موارد غیرقطعی در استثناها آمده‌اند"])
    summary_ws=wb[summary_name]
    for ridx,row in enumerate(result["summary"],2):
        for col,detail_sheet,title in [(19,sales_name,"مشاهده فروش"),(20,"ریز وصول ح.د","مشاهده وصول"),(21,"ریز وصول اسناد","مشاهده اسناد")]:
            target=first_rows.get(detail_sheet,{}).get((row[0],row[1]))
            if target: summary_ws.cell(ridx,col,title).hyperlink=f"#'{detail_sheet}'!A{target}"; summary_ws.cell(ridx,col).style="Hyperlink"
    output=io.BytesIO(); wb.save(output); return output.getvalue()


def build_workbook(result: dict) -> bytes:
    if xlsxwriter is None:
        return _build_workbook_openpyxl(result)
    output=io.BytesIO(); wb=xlsxwriter.Workbook(output,{"in_memory":True})
    header=wb.add_format({"bold":True,"font_color":"white","bg_color":"#173A5E","border":1,"align":"center","valign":"vcenter"})
    text_fmt=wb.add_format({"border":1,"align":"right","valign":"top"}); num_fmt=wb.add_format({"border":1,"num_format":"#,##0.000;[Red]-#,##0.000","align":"left"}); link_fmt=wb.add_format({"font_color":"#0563C1","underline":True,"border":1})
    month_fa=str(result["month"]).translate(str.maketrans("0123456789","۰۱۲۳۴۵۶۷۸۹")); summary_name="خلاصه ح.د و وصول"; sales_name=f"ریز ح.د فروش ماه {month_fa}"
    sheets=[
      (summary_name,["شرکت","تفصیل",f"مانده ح.د اول ماه {month_fa}",f"ح.د بدهکار فروش ماه {month_fa}",f"ح.د بستانکار فروش ماه {month_fa}",f"ح.د خالص فروش ماه {month_fa}","وصول مستقیم ح.د → بانک/واسط","ح.د → اسناد دریافتنی","ح.د → نقد/صندوق","سایر تسویه ح.د","جمع وصول مشتری (بانک + اسناد)","وصول از مانده قبل (FIFO)",f"وصول از فروش ماه {month_fa}",f"مانده ح.د فروش ماه {month_fa}","وصول اسناد → بانک","از وصول اسناد: قدیمی",f"از وصول اسناد: ماه {month_fa}","از وصول اسناد: منشأ نامشخص","لینک ریز فروش","لینک ریز وصول","لینک وصول اسناد"],result["summary"]),
      (sales_name,["شرکت","تفصیل","تاریخ","شماره سند/عطف","ردیف سند","بدهکار ح.د (میلیون تومان)","بستانکار ح.د (میلیون تومان)","حساب معین","شرح"],result["sales"]),
      ("ریز وصول ح.د",["شرکت","تفصیل","تاریخ","شماره سند/عطف","ردیف ح.د","نوع طرف","بستانکار ح.د","ردیف طرف","مبلغ تخصیص‌یافته","بدهکار طرف","حساب طرف","تفصیل طرف","شرح ح.د","شرح طرف","روش لینک","اختلاف"],result["receipts"]),
      ("ریز وصول اسناد",["شرکت","تفصیل مشتری","تاریخ وصول","شماره سند/عطف","ردیف اسناد","مبلغ وصول اسناد","ردیف بانک","بدهکار بانک","حساب اسناد","حساب بانک","شرح اسناد","شرح بانک","تاریخ دریافت اولیه","سند دریافت اولیه","مبنای شناسایی مشتری","دسته"],result["docs"]),
      ("استثناها",["نوع استثنا","شرکت","تفصیل","تاریخ","شماره سند/عطف","ردیف سند","مبلغ","حساب","شرح"],result["exceptions"])]
    first_rows={}
    for name,columns,rows in sheets:
        ws=wb.add_worksheet(name[:31]); ws.right_to_left(); ws.freeze_panes(1,2); ws.autofilter(0,0,max(1,len(rows)),len(columns)-1); ws.set_row(0,32); ws.set_column(0,len(columns)-1,16); ws.set_column(1,1,34); ws.set_column(max(0,len(columns)-2),len(columns)-1,44)
        for col,label in enumerate(columns): ws.write(0,col,label,header)
        first_rows[name]={}
        for ridx,row in enumerate(rows,1):
            if len(row)>1: first_rows[name].setdefault((row[0],row[1]),ridx+1)
            for col,value in enumerate(row):
                monetary=(name==summary_name and 2<=col<=17) or (name==sales_name and col in (5,6)) or (name=="ریز وصول ح.د" and col in (6,8,9,15)) or (name=="ریز وصول اسناد" and col in (5,7)) or (name=="استثناها" and col==6)
                ws.write(ridx,col,_to_million(value) if monetary else value,num_fmt if monetary else text_fmt)
    summary_ws=wb.get_worksheet_by_name(summary_name)
    for ridx,row in enumerate(result["summary"],1):
        key=(row[0],row[1])
        for col,detail_sheet,title in [(18,sales_name,"مشاهده فروش"),(19,"ریز وصول ح.د","مشاهده وصول"),(20,"ریز وصول اسناد","مشاهده اسناد")]:
            target=first_rows.get(detail_sheet,{}).get(key)
            if target: summary_ws.write_url(ridx,col,f"internal:'{detail_sheet}'!A{target}",link_fmt,title)
    controls=[["مانده خالص ح.د اول دوره",sum(r[2] for r in result["summary"]),"جمع مانده حساب‌های دریافتنی تجاری قبل از ماه"],["ح.د بدهکار سند فروش",sum(r[3] for r in result["summary"]),"ردیف‌های فکت دارای قرینه فروش/فاکتور"],["ح.د بستانکار سند فروش",sum(r[4] for r in result["summary"]),"برگشت/تعدیل در ردیف‌های فروش"],["ح.د خالص طرف فروش",sum(r[5] for r in result["summary"]),"بدهکار منهای بستانکار"],["وصول مستقیم بانک/واسط",sum(r[6] for r in result["summary"]),"تطبیق شرکت + تاریخ + شرح"],["ح.د → اسناد دریافتنی",sum(r[7] for r in result["summary"]),"دریافت چک از مشتری"],["ح.د → نقد/صندوق",sum(r[8] for r in result["summary"]),"وصول نقدی"],["سایر تسویه/انتقال ح.د",sum(r[9] for r in result["summary"]),"طرف غیر بانکی یا نامشخص"],["وصول اسناد → بانک",sum(r[14] for r in result["summary"]),"نقدشدن چک؛ وصول جدید مشتری نیست"],["تعداد استثناها",len(result["exceptions"]),"موارد غیرقطعی حذف نشده‌اند"]]
    ws=wb.add_worksheet("کنترل گزارش"); ws.right_to_left(); ws.set_column(0,0,34); ws.set_column(1,1,22); ws.set_column(2,2,58)
    for c,v in enumerate(["کنترل","مبلغ (میلیون تومان)","توضیح"]): ws.write(0,c,v,header)
    for r,row in enumerate(controls,1): ws.write(r,0,row[0],text_fmt); ws.write(r,1,_to_million(row[1]) if r<len(controls) else row[1],num_fmt); ws.write(r,2,row[2],text_fmt)
    ws=wb.add_worksheet("منبع گردش کل"); ws.right_to_left(); ws.set_column(0,2,48)
    for c,v in enumerate(["منبع","فایل منبع کامل","توضیح"]): ws.write(0,c,v,header)
    ws.write_row(1,0,["FactFinance","FactFinnance1.xlsx","فکت منبع؛ FinID به‌عنوان ردیف قابل ردیابی حفظ شده است"],text_fmt); ws.write_row(2,0,["محدودیت","شماره سند/عطف در فکت موجود نیست","شماره ساختگی تولید نشده و موارد غیرقطعی در استثناها آمده‌اند"],text_fmt)
    wb.close(); return output.getvalue()


def generate(input_path: str|Path,output_path: str|Path,year: int,month: int):
    result=analyze(load_source(input_path),year,month); Path(output_path).write_bytes(build_workbook(result)); return result


if __name__=="__main__":
    parser=argparse.ArgumentParser(description="Generate AR and collections report from FactFinnance1.xlsx"); parser.add_argument("--input",default="FactFinnance1.xlsx"); parser.add_argument("--output",default="AR_Collections_Report.xlsx"); parser.add_argument("--year",type=int,default=1405); parser.add_argument("--month",type=int,default=6); args=parser.parse_args(); result=generate(args.input,args.output,args.year,args.month); print(f"Created {args.output}: {len(result['summary'])} customers, {len(result['exceptions'])} exceptions")
