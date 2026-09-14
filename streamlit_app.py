import os
from pathlib import Path
import pandas as pd
import streamlit as st
from datetime import datetime
from report_generator import analyze, build_workbook, load_source

st.set_page_config(page_title="گزارش حساب‌های دریافتنی و وصول",page_icon="📊",layout="wide")
st.markdown("<style>html,body,[class*='css']{direction:rtl;text-align:right;font-family:Tahoma,sans-serif}[data-testid='stMetricValue']{direction:ltr;text-align:right}.block-container{padding-top:2rem}</style>",unsafe_allow_html=True)
st.title("گزارش حساب‌های دریافتنی و وصول"); st.caption(f"خروجی مطابق الگوی گزارش پنج شرکت · آخرین بروزرسانی برنامه: {datetime.now().strftime('%Y/%m/%d %H:%M')}")
with st.sidebar:
    st.header("تنظیمات گزارش"); uploaded=st.file_uploader("فکت جایگزین",type=["xlsx"],help="در صورت عدم انتخاب، فایل همراه پروژه خوانده می‌شود."); year=st.number_input("سال شمسی",1300,1500,1405); month=st.number_input("ماه",1,12,6); clicked=st.button("تولید گزارش",type="primary",use_container_width=True)
@st.cache_data(show_spinner=False)
def parse_default(path,modified): return load_source(path)
@st.cache_data(show_spinner=False)
def parse_upload(payload): return load_source(payload)
if "report" not in st.session_state: st.session_state.report=None
if clicked or st.session_state.report is None:
    with st.spinner("در حال خواندن فکت و ساخت گزارش…"):
        if uploaded: records=parse_upload(uploaded.getvalue())
        else:
            source=Path(os.environ.get("FACT_FINANCE_PATH",r"C:\Users\a.farshchian\Desktop\AR_Aging_Report\FactFinnance1.xlsx"))
            if not source.exists(): st.error("فایل فکت محلی پیدا نشد؛ مسیر FACT_FINANCE_PATH را تنظیم یا فایل را آپلود کنید."); st.stop()
            records=parse_default(str(source),source.stat().st_mtime)
        st.session_state.report=analyze(records,int(year),int(month))
report=st.session_state.report
columns=["شرکت","تفصیل","مانده اول دوره","بدهکار فروش","بستانکار فروش","خالص فروش","وصول بانک/واسط","دریافت چک","نقد/صندوق","سایر تسویه","جمع وصول بانک و چک","وصول فروش قبلی","وصول فروش جاری","مانده فروش ماه","وصول اسناد به بانک","چک قدیمی","چک جاری","منشأ نامشخص","لینک فروش","لینک وصول","لینک اسناد"]
summary=pd.DataFrame(report["summary"],columns=columns); companies=sorted(summary["شرکت"].dropna().astype(str).unique()) if not summary.empty else []; selected=st.multiselect("شرکت",companies,default=companies); query=st.text_input("جست‌وجوی مشتری"); view=summary.copy()
if selected: view=view[view["شرکت"].astype(str).isin(selected)]
if query: view=view[view["تفصیل"].astype(str).str.contains(query,case=False,na=False)]
metrics=[("خالص فروش",view["خالص فروش"].sum()),("وصول بانک/واسط",view["وصول بانک/واسط"].sum()),("دریافت چک",view["دریافت چک"].sum()),("جمع وصول",view["جمع وصول بانک و چک"].sum()),("مانده فروش ماه",view["مانده فروش ماه"].sum()),("استثناهای باز",len(report["exceptions"]))]
for col,(label,value) in zip(st.columns(6),metrics): col.metric(label,f"{value/10_000_000:,.1f}" if label!="استثناهای باز" else f"{value:,}")
st.caption("مبالغ: میلیون تومان؛ محاسبات داخلی: ریال دقیق منبع"); st.subheader("خلاصه مشتریان"); display=view.drop(columns=["لینک فروش","لینک وصول","لینک اسناد"],errors="ignore").copy()
for column in display.columns[2:]: display[column]=display[column]/10_000_000
st.dataframe(display,use_container_width=True,height=460,hide_index=True)
st.download_button("دریافت فایل Excel مطابق الگو",data=build_workbook(report),file_name=f"AR_Collections_{int(year)}_{int(month):02d}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
with st.expander("محدودیت منبع"): st.write("فکت شماره سند/عطف مستقل ندارد. FinID به‌عنوان ردیف منبع حفظ شده و هیچ شماره سند ساختگی تولید نمی‌شود؛ لینک‌های غیرقطعی در شیت استثناها گزارش می‌شوند.")
