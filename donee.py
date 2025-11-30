# -*- coding: utf-8 -*-
"""
Streamlit app: Agent cards + Supervisor report (Overview + Daily) — HTML only

- Agent cards (single + bulk-by-city HTML) → labels فارسی
- Supervisor report (Overview + Daily) → هدر + جدول کلی هفته + جدول روزانه
- ادغام «نمره آموزش» و «نمره حضور» → «نمره QC راننده‌ها»
- رنگ‌بندی ردیف‌های روزانهٔ سرپرست بر اساس Tier: آبی (1)، سبز (2)، نارنجی (3)
- Batch download کارت‌ها (N-by-N)
- «روزهای بدون مشارکت» = absence
- ستون ACQ → «تعداد جذب» در کارت بازاریاب
- Top 10 همه‌جانبه در هر شهر (HTML)
- رتبهٔ هر بازاریاب در شهر خودش (X از Y) روی کارت
- NEW: Badge 🔴 برای «روزهای بدون مشارکت > ۲۰٪» در Overview
- NEW: نمودار Altair ترند ۴ هفته‌ای متریک‌ها زیر کارت تکی
"""

import re
import math
import base64
from datetime import datetime, timedelta
import streamlit as st
import pandas as pd

# --- Optional: Altair chart ---
try:
    import altair as alt
    _ALT_OK = True
except Exception:
    _ALT_OK = False

# ---------- Helpers ----------
def spacer(px: int = 8):
    st.markdown(f"<div style='height:{px}px'></div>", unsafe_allow_html=True)

def add_badges_overview(df: pd.DataFrame) -> pd.DataFrame:
    """فقط ایموجی 🔴 کنار «روزهای بدون مشارکت» اگر بالای ۲۰٪ باشد."""
    if df is None or getattr(df, "empty", True):
        return df
    out = df.copy()

    def _to_pct_float(v):
        if v is None: return None
        s = str(v).strip()
        if s.endswith('%'): s = s[:-1]
        s = s.replace(',', '').replace('٪','')
        try: return float(s)
        except: return None

    col = "روزهای بدون مشارکت"
    if col in out.columns:
        out[col] = out[col].apply(
            lambda x: f"{x} 🔴" if ((_to_pct_float(x) is not None) and (_to_pct_float(x) > 20.0)) else x
        )
    return out

# ---------- Utils ----------
def normalize_header(h: str) -> str:
    if h is None: return ""
    s = str(h).replace('"','').replace("'","").replace("\t"," ").replace("\n"," ")
    # underscore و # هم پاک/به فاصله تبدیل می‌شن
    s = s.replace("_", " ").replace("#", "")
    s = re.sub(r"[,:;؛،\-\(\)\[\]\{\}\\/]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace("ي","ی").replace("ك","ک")
    return s.lower()


def fmt_val(x, percent=False, nd=2):
    """Formatter: show '-' if empty/NaN, else numeric/percent."""
    try:
        if x is None: return "-"
        if isinstance(x, str) and x.strip() == "": return "-"
        if hasattr(pd, "isna") and pd.isna(x): return "-"
    except Exception:
        pass
    try:
        v = float(x)
    except Exception:
        return str(x)
    if percent:
        if -1.01 <= v <= 1.01: v *= 100.0
        return f"{v:.0f}%"
    if abs(v - int(v)) < 1e-9:
        return f"{int(v)}"
    return f"{v:.{nd}f}"

_PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")

def to_numeric_clean(series):
    """Convert a pandas Series to numeric after stripping commas and فارسی digits."""
    if series is None:
        return pd.Series(dtype="float64")
    return pd.to_numeric(
        (
            series.astype(str)
                  .str.replace("٬", "", regex=False)
                  .str.replace(",", "", regex=False)
                  .str.translate(_PERSIAN_DIGITS)
                  .str.strip()
                  .replace({"": None})
        ),
        errors="coerce",
    )

# ستون‌های نقص مدارک (از فایل tatbigh)
DOC_ISSUE_COLS = [
    "عکس",
    "بیمه",
    "ویدئو",
    "قرارداد",
    "مغایرت",
    "ثبت اشتباه",
    "فاقد مدارک",
    "کارت ماشین",
    "معاینه فنی",
    "گواهینامه",
    "وضوح",
    "4 طرف خودرو",
]

def load_excel(file):
    return pd.read_excel(file)

# === Gregorian → Jalali (no external libs) ===
def _is_gregorian_leap(y: int) -> int:
    return 1 if (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0) else 0

def gregorian_to_jalali(g_y: int, g_m: int, g_d: int):
    g_days_in_month = [31, 28 + _is_gregorian_leap(g_y), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    j_days_in_month = [31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]

    gy = g_y - 1600
    gm = g_m - 1
    gd = g_d - 1

    g_day_no = 365 * gy + (gy + 3) // 4 - (gy + 99) // 100 + (gy + 399) // 400
    for i in range(gm): g_day_no += g_days_in_month[i]
    g_day_no += gd

    j_day_no = g_day_no - 79
    j_np = j_day_no // 12053
    j_day_no %= 12053

    jy = 979 + 33 * j_np + 4 * (j_day_no // 1461)
    j_day_no %= 1461

    if j_day_no >= 366:
        jy += (j_day_no - 1) // 365
        j_day_no = (j_day_no - 1) % 365

    jm = 0
    while jm < 11 and j_day_no >= j_days_in_month[jm]:
        j_day_no -= j_days_in_month[jm]; jm += 1
    jd = j_day_no + 1
    return jy, jm + 1, jd

PERSIAN_MONTHS = [
    "فروردین","اردیبهشت","خرداد","تیر","مرداد","شهریور",
    "مهر","آبان","آذر","دی","بهمن","اسفند"
]

def to_jalali_words(dt, persian_digits=True) -> str:
    """خروجی: '۳ مهر ۱۴۰۴'."""
    if isinstance(dt, pd.Timestamp): dt = dt.to_pydatetime().date()
    if isinstance(dt, datetime): dt = dt.date()
    y, m, d = dt.year, dt.month, dt.day
    jy, jm, jd = gregorian_to_jalali(y, m, d)
    out = f"{jd} {PERSIAN_MONTHS[jm-1]} {jy}"
    if persian_digits:
        out = out.translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))
    return out

def to_jalali_day_month(dt, persian_digits=True) -> str:
    """خروجی: '۳ مهر' (فقط روز و ماه، بدون سال)."""
    if isinstance(dt, pd.Timestamp): dt = dt.to_pydatetime().date()
    if isinstance(dt, datetime): dt = dt.date()
    y, m, d = dt.year, dt.month, dt.day
    jy, jm, jd = gregorian_to_jalali(y, m, d)
    # ترتیب: عدد اول، بعد ماه
    out = f"{jd} {PERSIAN_MONTHS[jm-1]}"
    if persian_digits:
        out = out.translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))
    return out

def export_qc_last_record_for_month(df, mp, jalali_month, year=1404, filename="output.xlsx"):
    """
    ساخت خروجی اکسل برای آخرین TOTAL QC SCORE هر بازاریاب در یک ماه شمسی مشخص
    jalali_month = 7 (مهر) یا 8 (آبان)
    year = سال شمسی (پیش‌فرض 1404)
    """

    date_col = mp["date"]
    id_col   = mp["id"]

    tmp = df.copy()
    tmp["_date"] = pd.to_datetime(tmp[date_col], errors="coerce")

    # تبدیل میلادی → شمسی
    def _jalali_parts(dt):
        """
        تبدیل یک تاریخ میلادی (datetime) به مؤلفه‌های شمسی (سال، ماه، روز).
        اگر مقدار ورودی NaT یا None باشد، یک سه‌تایی از None برمی‌گرداند تا در فراخوانی‌های بعدی
        (مانند _jalali_parts(x)[0]) خطایی به وجود نیاید.
        """
        if pd.isna(dt):
            # مقدار datetime نامعتبر یا خالی است؛ بازگشت یک سه‌تایی None
            return (None, None, None)
        y, m, d = gregorian_to_jalali(dt.year, dt.month, dt.day)
        return (y, m, d)

    tmp["_jy"] = tmp["_date"].apply(lambda x: _jalali_parts(x)[0] if x is not None else None)
    tmp["_jm"] = tmp["_date"].apply(lambda x: _jalali_parts(x)[1] if x is not None else None)

    # فیلتر روی ماه مورد نظر
    tmp = tmp[(tmp["_jy"] == year) & (tmp["_jm"] == jalali_month)]
    if tmp.empty:
        print(f"هیچ داده‌ای برای ماه {jalali_month} پیدا نشد.")
        return None

    # اگر ستون total qc score وجود ندارد → محاسبه کن
    if "TOTAL QC SCORE" not in tmp.columns:
        tmp["TOTAL QC SCORE"] = tmp.apply(compute_total_qc_score_from_row, axis=1)

    # آخرین رکورد هر آیدی
    tmp = tmp.sort_values("_date")
    last_rows = tmp.groupby(id_col).tail(1)

    # انتخاب ستون‌های لازم
    out = last_rows[[id_col, mp.get("name",""), mp.get("city",""), "TOTAL QC SCORE"]].copy()
    out = out.rename(columns={mp.get("name",""): "Name", mp.get("city",""): "City"})

    # ذخیره در اکسل
    out.to_excel(filename, index=False)
    print(f"فایل '{filename}' ساخته شد.")

    return out

def find_column_mapping(df: pd.DataFrame) -> dict:
    m, norm = {}, {c: normalize_header(c) for c in df.columns}
    def find(*keys):
        for col, n in norm.items():
            for k in keys:
                if k in n:
                    return col
        return None

    # raw columns
    m["date"]             = find("miladi", "date")   # ← دیگه فقط miladi نیست
    m["id"]               = find("آیدی","ایدی","id")
    m["name"]             = find("نام و نام خانوادگی","نام","name","fullname","full name")
    m["city"]             = find("city","شهر")
    m["supervisor"]       = find("سرپرست","supervisor")
    m["teamlead"]         = find("teamlead","تیم لید","تیم‌ لید")
    m["tier"]             = find("tier","تیرب","تییر")
    m["week"] = find(
        "هفته شمسی", 
        "week shamsi",
        "week_shamsi",
        "هفته", 
        "week", 
        "weeknum", 
        "wk"
    )

    m["team_size_excel"]  = find("teamsize")                    # TeamSize
    m["actual_ts_excel"]  = find("actualteamsize","actual team size")  # actualTeamSize

    m["participants"]     = find("participants","مشارکت")
    m["absence"]          = find("absence","no participation","روزهای بدون مشارکت")
    m["goldentime"]       = find("goldentime","golden time","ساعات طلایی")
    m["fake_sub"]         = find("offakesubmissions","fake submission","fake_sub","fake submissions")

    m["total_hours"]      = find("totalworkhours","total work hours")
    m["shift_delay"]      = find("shift delay","shift_delay","با تاخیر")
    m["false_check"]      = find("false check","false_check","خوداظهاری اشتباه")
    m["incomplete"]       = find("incomplete","ناقص")

    # نمره‌های QC راننده‌ها
    m["edu_score"]        = find("eduction score","education score","eduction","eduction_score")
    m["presence_score"]   = find("presence score","presence_score")

    # نمره‌های QC ناظر میدانی / بایکر
    m["presence_q_daily"] = find(
        "presencequality",       # الگوی کلی
        "presence quality",
        "presencequality daily",
        "presencequality_daily", # اگه جایی underline بوده
        "presence quality biker"
    )

    m["banner_q_daily"] = find(
        "bannerquality",
        "banner quality",
        "bannerquality daily",
        "bannerquality_daily"
    )

    # برای سازگاری با کدهای قدیمی که presence_qual را استفاده می‌کردند
    m["presence_qual"] = m.get("presence_q_daily") or m.get("banner_q_daily")
    m["banner_daily"] = find("banner_daily", "banner daily")
    m["active_drivers"] = find("active_drivers", "active drivers", "راننده فعال", "راننده‌های فعال")
    # Activation و Tatbigh
    m["activation"]       = find("activation","فعال سازی","فعالسازی")
    m["tatbigh"]          = find("tatbigh","تطبیق")

    # ACQ (acquisitions per agent) — الان daily_acq
    m["acq"]              = find("daily acq","daily_acq","acq","تعداد جذب")

    # sanity
    if not m.get("id"):
        raise ValueError("ستون «آیدی» پیدا نشد.")
    if not m.get("date"):
        raise ValueError("ستون تاریخ (date / miladi) پیدا نشد.")
    if not m.get("actual_ts_excel"):
        raise ValueError("ستون «actualTeamSize» پیدا نشد.")
    if not m.get("team_size_excel"):
        raise ValueError("ستون «TeamSize» پیدا نشد.")

    # supervisor semantics:
    m["assigned_col"] = m["actual_ts_excel"]  # تخصیص داده شده
    m["people_col"]   = m["team_size_excel"]  # افراد تیم
    return m


def filter_last_7_days(df: pd.DataFrame, date_col: str, anchor_date: datetime = None) -> pd.DataFrame:
    tmp = df.copy()
    tmp["_date"] = pd.to_datetime(tmp[date_col], errors="coerce").dt.date
    if anchor_date is None: anchor_date = pd.to_datetime(tmp["_date"]).max()
    if pd.isna(anchor_date): return tmp.iloc[0:0]
    anchor = pd.to_datetime(anchor_date).date()
    start = anchor - timedelta(days=6)
    return tmp[(tmp["_date"] >= start) & (tmp["_date"] <= anchor)]

def filter_last_7_days_per_group(df: pd.DataFrame, date_col: str, group_col: str) -> pd.DataFrame:
    tmp = df.copy()
    tmp["_date"] = pd.to_datetime(tmp[date_col], errors="coerce").dt.date
    # آخرین تاریخ در هر شهر
    grp_max = tmp.groupby(group_col)["_date"].transform("max")
    # بازهٔ ۷ روزه نسبت به آخرین تاریخ همان شهر
    mask = (tmp["_date"] >= (grp_max - pd.to_timedelta(6, unit="D"))) & (tmp["_date"] <= grp_max)
    return tmp[mask]

# ---------- Agent weekly (cards) ----------
def _combine_edu_presence_label(row) -> str:
    """
    QC راننده‌ها از میانگین Presence/Education:
    """
    def parse_num(v):
        if v is None or (hasattr(pd,"isna") and pd.isna(v)): return None
        s = str(v).strip().replace(",", "")
        if s.endswith("%"): s = s[:-1].strip()
        if s == "": return None
        try:
            x = float(s)
            if x <= 1.01: x *= 100.0
            return x
        except Exception:
            return None

    vals = []
    for c in row.index:
        n = normalize_header(c)
        if (("presence" in n and "score" in n) or
            ("education" in n and "score" in n) or
            ("eduction" in n and "score" in n) or
            ("نمره" in c and ("حضور" in c or "آموزش" in c))):
            pv = parse_num(row[c])
            if pv is not None: vals.append(pv)

    if not vals:
        return "متوسط"
    avg = sum(vals)/len(vals)
    if avg <= 30:
        return "ضعیف"
    elif avg <= 70:
        return "خوب"
    else:
        return "عالی"

def compute_weekly_stats(df: pd.DataFrame, mp: dict) -> pd.DataFrame:
    id_col = mp["id"]

    # یک کپی از df برای اینکه ستون موقتی اضافه کنیم
    df = df.copy()

    # ------ ساخت نمره QC ناظران میدانی از روی دو ستون روزانه ------
    pq_col = mp.get("presence_q_daily")
    bq_col = mp.get("banner_q_daily")


    def _parse_pct(v):
        if v is None or (hasattr(pd,"isna") and pd.isna(v)):
            return None
        s = str(v).strip().replace(",", "")
        if s.endswith("%") or s.endswith("٪"):
            s = s[:-1].strip()
        if s == "":
            return None
        try:
            x = float(s)
            # اگر ۰–۱ بود، ببریم روی ۰–۱۰۰
            if x <= 1.01:
                x *= 100.0
            return x
        except Exception:
            return None

    if pq_col or bq_col:
        def _combine_field_q(row):
            vals = []
            if pq_col and pq_col in row.index:
                v = _parse_pct(row[pq_col])
                if v is not None:
                    vals.append(v)
            if bq_col and bq_col in row.index:
                v = _parse_pct(row[bq_col])
                if v is not None:
                    vals.append(v)
            if not vals:
                return None
            return sum(vals)/len(vals)

        df["_field_q"] = df.apply(_combine_field_q, axis=1)
    # -------------------------------------------------------------

    # روزهای کاری
    unique_days = (df[[id_col,"_date"]].dropna()
                   .drop_duplicates([id_col,"_date"])
                   .groupby(id_col)["_date"].nunique()
                   .rename("روزهای کاری"))

    agg = {}
    for k in ["name","city","supervisor","teamlead"]:
        col = mp.get(k)
        if col: agg[col] = "first"

    if pq_col:
        agg[pq_col] = "mean"

    if bq_col:
        agg[bq_col] = "mean"

    # جمع‌ها (Agent)
    if mp.get("team_size_excel"):
        agg[mp["team_size_excel"]] = "sum"  # تعداد فیلدها
    if mp.get("absence"):      agg[mp["absence"]]      = "sum"  # روزهای بدون مشارکت
    if mp.get("incomplete"):   agg[mp["incomplete"]]   = "sum"
    if mp.get("fake_sub"):     agg[mp["fake_sub"]]     = "sum"
    if mp.get("shift_delay"):  agg[mp["shift_delay"]]  = "sum"
    if mp.get("false_check"):  agg[mp["false_check"]]  = "sum"
    if mp.get("goldentime"):   agg[mp["goldentime"]]   = "sum"   # ✅ اضافه شد
    if mp.get("total_hours"):  agg[mp["total_hours"]]  = "sum"   # ✅ اضافه شد
    if mp.get("acq"):          agg[mp["acq"]]          = "sum"   # ACQ
    if mp.get("active_drivers"):   agg[mp["active_drivers"]]   = "sum" 
    if mp.get("tatbigh"):      agg[mp["tatbigh"]]      = "mean"  

    for col in DOC_ISSUE_COLS:
        if col in df.columns:
            agg[col] = "sum"
    if "تعداد راننده فعال" in df.columns:
        agg["تعداد راننده فعال"] = "sum"


    # میانگین‌ها
    for k in ["edu_score","presence_score","presence_qual"]:
        col = mp.get(k)
        if col: agg[col] = "mean"

    # ✅ ستون‌های جزئی فرم ارزیابی راننده (۰–۱)
    detail_cols = [
        "مراجعه",
        "وضعیت حضور",
        "پوشش و رفتار مناسب",
        "مدت زمان ثبت نام",
        "مزایا",
        "تسویه روزانه",
        "پشتیبانی",
        "وعده ی غیرواقعی",
        "وعده غیرواقعی",
    ]
    for col in df.columns:
        if str(col).strip() in detail_cols:
            agg[col] = "mean"

    # میانگین نمره QC ناظران میدانی از ستون موقت
    if "_field_q" in df.columns:
        agg["_field_q"] = "mean"

    base = df.groupby(id_col).agg(agg).reset_index().merge(unique_days, left_on=id_col, right_index=True, how="left")

    # برچسب‌های فارسی کارت بازاریاب
    rename_map = {id_col:"آیدی"}
    if mp.get("name"):           rename_map[mp["name"]] = "نام و نام خانوادگی"
    if mp.get("city"):           rename_map[mp["city"]] = "شهر"
    if mp.get("supervisor"):     rename_map[mp["supervisor"]] = "سرپرست"
    if mp.get("teamlead"):       rename_map[mp["teamlead"]] = "تیم‌ لید"

    if mp.get("team_size_excel"): rename_map[mp["team_size_excel"]] = "تعداد فیلد های تخصیص داده شده"
    if mp.get("absence"):         rename_map[mp["absence"]]         = "روزهای بدون مشارکت"
    if mp.get("incomplete"):      rename_map[mp["incomplete"]]      = "مشارکت‌های ناقص"
    if mp.get("fake_sub"):        rename_map[mp["fake_sub"]]        = "مشارکت خارج از محدوده مجاز"
    if mp.get("shift_delay"):     rename_map[mp["shift_delay"]]     = "مشارکت با تاخیر"
    if mp.get("false_check"):     rename_map[mp["false_check"]]     = "خوداظهاری اشتباه"
    if mp.get("goldentime"):      rename_map[mp["goldentime"]]      = "ساعات طلایی"        
    if mp.get("total_hours"):     rename_map[mp["total_hours"]]     = "ساعات کاری کل"     
    if mp.get("acq"):             rename_map[mp["acq"]]             = "تعداد جذب"
    if mp.get("activation"):      rename_map[mp["activation"]]      = "فعال سازی"
    if mp.get("tatbigh"):         rename_map[mp["tatbigh"]]         = "تطبیق"
    if mp.get("active_drivers"):  rename_map[mp["active_drivers"]]  = "تعداد راننده‌های فعال"



    if mp.get("edu_score"):       rename_map[mp["edu_score"]]       = "نمره آموزش توسط راننده ها"
    if mp.get("presence_score"):  rename_map[mp["presence_score"]]  = "نمره حضور توسط راننده ها"

    # نمره QC ناظران میدانی از ستون موقت
    if "_field_q" in base.columns:
        rename_map["_field_q"] = "نمره QC ناظران میدانی"
    # نام ستون‌های جزئی را همان نگه می‌داریم (برای بولت‌پوینت‌ها)
    for dc in [
        "مراجعه",
        "وضعیت حضور",
        "پوشش و رفتار مناسب",
        "مدت زمان ثبت نام",
        "مزایا",
        "تسویه روزانه",
        "پشتیبانی",
        "وعده ی غیرواقعی",
        "وعده غیرواقعی",
    ]:
        if dc in base.columns:
            rename_map[dc] = dc

    out = base.rename(columns=rename_map)


    if "تعداد راننده‌های فعال" in out.columns and "تعداد جذب" in out.columns:
        active_num = to_numeric_clean(out["تعداد راننده‌های فعال"])
        acq_num = to_numeric_clean(out["تعداد جذب"])

        out["درصد راننده‌های فعال"] = (active_num / acq_num.replace(0, pd.NA)) * 100

    doc_cols_present = [c for c in DOC_ISSUE_COLS if c in out.columns]
    if doc_cols_present:
        doc_numeric = out[doc_cols_present].apply(lambda col: to_numeric_clean(col))
        out["تعداد نقص مدارک"] = doc_numeric.sum(axis=1, min_count=1)

    if "تطبیق" in out.columns and "درصد نقص مدارک" not in out.columns:
        out["درصد نقص مدارک"] = out["تطبیق"]
    if "فعال سازی" in out.columns and "درصد راننده‌های فعال" not in out.columns:
        out["درصد راننده‌های فعال"] = out["فعال سازی"]

    # ادغام دو نمره → نمره QC راننده‌ها
    out["نمره QC راننده ها"] = out.apply(_combine_edu_presence_label, axis=1)

    # مقادیر عددی
    numeric_cols = [
        "تعداد فیلد های تخصیص داده شده","روزهای بدون مشارکت","مشارکت‌های ناقص",
        "مشارکت خارج از محدوده مجاز","مشارکت با تاخیر","خوداظهاری اشتباه",
        "روزهای کاری","تعداد جذب","نمره QC ناظران میدانی","فعال سازی","تطبیق",
        "ساعات طلایی","ساعات کاری کل","تعداد راننده‌های فعال","تعداد نقص مدارک",
    ]
    numeric_cols += [c for c in DOC_ISSUE_COLS if c in out.columns]
    for c in numeric_cols:
        if c in out.columns:
            out[c] = to_numeric_clean(out[c])

    # ✅ محاسبه TOTAL QC SCORE بر اساس متریک‌های هفتگی (بعد از تبدیل به عدد)
    if all(col in out.columns for col in [
        "خوداظهاری اشتباه",
        "ساعات طلایی",
        "ساعات کاری کل",
        "روزهای کاری",
        "مشارکت با تاخیر",
        "مشارکت‌های ناقص",
        "تعداد فیلد های تخصیص داده شده",
    ]):
        out["TOTAL QC SCORE"] = out.apply(compute_total_qc_score_from_row, axis=1)

    return out



# ---------- Weekly 4-week trend (for chart) ----------
def compute_weekly_rates_for_agent(df: pd.DataFrame, mp: dict, agent_id) -> pd.DataFrame:
    """
    خروجی: DataFrame با ستون‌های [week, metric, value] فقط برای همان آیدی، ۴ هفتهٔ اخیر.
    متریک‌ها را به‌صورت «درصد نسبت به assigned» محاسبه می‌کنیم.
    """
    if not mp.get("week"):
        return pd.DataFrame()

    need = [mp["id"], mp["week"], mp["assigned_col"]]
    for k in ["absence","incomplete","shift_delay","goldentime","fake_sub","false_check"]:
        if mp.get(k):
            need.append(mp[k])

    sub = df[need].copy()

    # هفته را عددی و فیلتر ۴ هفته‌ی آخر
    sub["week"] = pd.to_numeric(sub[mp["week"]], errors="coerce")
    last_weeks = sorted(sub["week"].dropna().unique())[-4:]
    sub = sub[sub["week"].isin(last_weeks)]

    # فقط همان آیدی
    sub = sub[sub[mp["id"]].astype(str) == str(agent_id)]
    if sub.empty:
        return pd.DataFrame(columns=["week","metric","value"])

    # تجمیع هفتگی
    g = sub.groupby("week").agg(asg_sum=(mp["assigned_col"], "sum")).reset_index()

    def _sum_col(key, alias):
        if mp.get(key):
            s = sub.groupby("week")[[mp[key]]].sum().rename(columns={mp[key]: alias})
            return s
        return None

    for key, alias in [
        ("absence","abs_sum"),
        ("incomplete","inc_sum"),
        ("shift_delay","sdl_sum"),
        ("goldentime","gld_sum"),
        ("fake_sub","fk_sum"),
        ("false_check","fc_sum"),
    ]:
        s = _sum_col(key, alias)
        if s is not None:
            g = g.merge(s, on="week", how="left")

    # عددی‌سازی
    for c in g.columns:
        if c != "week":
            g[c] = pd.to_numeric(g[c], errors="coerce")
    den = g["asg_sum"].replace(0, pd.NA)

    rows = []
    def _add(fa_label, sum_col):
        if sum_col in g.columns:
            vals = (g[sum_col] / den) * 100
            rows.append(pd.DataFrame({
                "week": g["week"],
                "metric": [fa_label]*len(g),
                "value": vals.astype(float)
            }))

    _add("روزهای بدون مشارکت", "abs_sum")
    _add("مشارکت‌های ناقص", "inc_sum")
    _add("مشارکت با تاخیر", "sdl_sum")
    _add("ساعات طلایی شیفت", "gld_sum")
    _add("مشارکت خارج از محدوده مجاز", "fk_sum")
    _add("خوداظهاری اشتباه", "fc_sum")

    if not rows:
        return pd.DataFrame(columns=["week","metric","value"])
    out = pd.concat(rows, ignore_index=True)
    out = out[out["week"].isin(last_weeks)].sort_values("week")
    return out

def compute_weekly_rates_for_city(df: pd.DataFrame, mp: dict, city_name) -> pd.DataFrame:
    if not mp.get("week") or not mp.get("city"):
        return pd.DataFrame(columns=["week","metric","value"])

    need = [mp["city"], mp["week"], mp["assigned_col"]]
    for k in ["absence","incomplete","shift_delay","goldentime","fake_sub","false_check"]:
        if mp.get(k):
            need.append(mp[k])

    sub = df[need].copy()
    sub = sub[sub[mp["city"]].astype(str) == str(city_name)]

    sub["week"] = pd.to_numeric(sub[mp["week"]], errors="coerce")
    last_weeks = sorted(sub["week"].dropna().unique())[-4:]
    sub = sub[sub["week"].isin(last_weeks)]
    if sub.empty:
        return pd.DataFrame(columns=["week","metric","value"])

    g = sub.groupby("week").agg(asg_sum=(mp["assigned_col"], "sum")).reset_index()

    def _sum_col(key, alias):
        if mp.get(key):
            s = sub.groupby("week")[[mp[key]]].sum().rename(columns={mp[key]: alias})
            return s
        return None

    for key, alias in [
        ("absence","abs_sum"),
        ("incomplete","inc_sum"),
        ("shift_delay","sdl_sum"),
        ("goldentime","gld_sum"),
        ("fake_sub","fk_sum"),
        ("false_check","fc_sum"),
    ]:
        s = _sum_col(key, alias)
        if s is not None:
            g = g.merge(s, on="week", how="left")

    for c in g.columns:
        if c != "week":
            g[c] = pd.to_numeric(g[c], errors="coerce")
    den = g["asg_sum"].replace(0, pd.NA)

    rows = []
    def _add(fa_label, sum_col):
        if sum_col in g.columns:
            vals = (g[sum_col] / den) * 100
            rows.append(pd.DataFrame({
                "week": g["week"],
                "metric": [fa_label]*len(g),
                "value": vals.astype(float)
            }))

    _add("روزهای بدون مشارکت", "abs_sum")
    _add("مشارکت‌های ناقص", "inc_sum")
    _add("مشارکت با تاخیر", "sdl_sum")
    _add("ساعات طلایی شیفت", "gld_sum")
    _add("مشارکت خارج از محدوده مجاز", "fk_sum")
    _add("خوداظهاری اشتباه", "fc_sum")

    if not rows:
        return pd.DataFrame(columns=["week","metric","value"])

    out = pd.concat(rows, ignore_index=True)
    out = out[out["week"].isin(last_weeks)].sort_values("week")
    return out

# ======= NEW: last-4-weeks ACQ helpers =======
def build_acq_trend_last4weeks(df: pd.DataFrame, mp: dict, agent_id):
    """ترند جذب ۴ هفته اخیر بر اساس ستون هفته شمسی و daily_acq"""
    if not mp.get("week") or not mp.get("acq"):
        return None
    
    try:
        # ستون‌های لازم (شامل _date برای محاسبه بازه تاریخ)
        cols_needed = [mp["id"], mp["week"], mp["acq"]]
        if "_date" in df.columns:
            cols_needed.append("_date")
        df2 = df[cols_needed].copy()

        # فقط همین آیدی (با strip برای اطمینان از تطابق)
        df2 = df2[df2[mp["id"]].astype(str).str.strip() == str(agent_id).strip()]

        if df2.empty:
            return None

        # عددی‌سازی (با نام‌های موقت برای جلوگیری از تداخل)
        df2["_week_num"] = to_numeric_clean(df2[mp["week"]])
        df2["_acq_num"]  = to_numeric_clean(df2[mp["acq"]])

        # حذف ردیف‌هایی که week معتبر ندارند
        df2 = df2.dropna(subset=["_week_num"])
        if df2.empty:
            return None

        # جایگزین کردن NaN در acq با 0 (اگر acq خالی بود)
        df2["_acq_num"] = df2["_acq_num"].fillna(0)

        # همه هفته‌ها
        weeks_sorted = sorted(df2["_week_num"].unique())

        if not weeks_sorted:
            return None

        # آخرین ۴ هفته
        last4 = weeks_sorted[-4:]

        df2 = df2[df2["_week_num"].isin(last4)]

        if df2.empty:
            return None

        # جمع جذب هر هفته و محاسبه بازه تاریخ
        if "_date" in df2.columns:
            # تبدیل _date به date اگر datetime است
            df2["_date_clean"] = pd.to_datetime(df2["_date"], errors="coerce").dt.date
            # محاسبه min و max تاریخ برای هر هفته
            agg = df2.groupby("_week_num").agg({
                "_acq_num": "sum",
                "_date_clean": ["min", "max"]
            }).reset_index()
            agg.columns = ["week", "acq", "date_min", "date_max"]
            
            # ساخت برچسب بازه تاریخ برای هر هفته با فرمت "X [ماه] تا" و "Y [ماه]" (بدون سال)
            def make_date_label(d_min, d_max):
                if pd.isna(d_min) or pd.isna(d_max):
                    return None
                try:
                    from_str = to_jalali_day_month(d_min, persian_digits=True)  # فقط روز و ماه
                    to_str = to_jalali_day_month(d_max, persian_digits=True)  # فقط روز و ماه
                    # برگرداندن یک رشته با جداکننده خاص برای تقسیم بعدی
                    return f"START:{from_str}|END:{to_str}"
                except:
                    return None
            
            agg["label"] = agg.apply(lambda row: make_date_label(row["date_min"], row["date_max"]), axis=1)
            # اگر label خالی شد، fallback به شماره هفته
            for idx in range(len(agg)):
                if pd.isna(agg.iloc[idx]["label"]) or agg.iloc[idx]["label"] is None:
                    agg.iloc[idx, agg.columns.get_loc("label")] = f"هفته {idx + 1}"
            agg = agg[["week", "acq", "label"]]
        else:
            # اگر _date نداشتیم، همان روش قبلی
            agg = df2.groupby("_week_num")["_acq_num"].sum().reset_index()
            agg = agg.rename(columns={"_week_num": "week", "_acq_num": "acq"})
            agg = agg.sort_values("week").reset_index(drop=True)
            # برچسب‌ها: هفته ۱..۴
            agg["label"] = [f"هفته {i+1}" for i in range(len(agg))]

        agg = agg.sort_values("week").reset_index(drop=True)
        return agg
    except Exception:
        # در صورت خطا، None برگردان
        return None

def build_acq_trend_svg_from_df(agg_df, width=420, height=160, point_fs=14):
    if agg_df is None or agg_df.empty:
        return None

    weeks = agg_df["label"].tolist()
    vals  = agg_df["acq"].astype(float).tolist()

    W, H = width, height
    pad  = 28
    top_pad = 22
    label_h = 40  # افزایش ارتفاع برای تاریخ‌های طولانی‌تر

    plot_h = H - (pad + label_h + top_pad)
    plot_y0 = top_pad
    plot_y1 = top_pad + plot_h

    n = len(weeks)
    if n == 1:
        xs = [W/2]
    else:
        step = (W - 2*pad) / (n - 1)
        xs = [pad + i*step for i in range(n)]

    vmin, vmax = min(vals), max(vals)
    if vmin == vmax:
        vmin -= 1; vmax += 1

    def y_map(v):
        return plot_y1 - ((v - vmin) / (vmax - vmin)) * plot_h

    ys = [y_map(v) for v in vals]

    bg = f'<rect x="0" y="0" width="{W}" height="{H}" fill="#ffffff"/>'

    grid = ""
    for gy in [plot_y0, (plot_y0+plot_y1)/2, plot_y1]:
        grid += f'<line x1="{pad}" y1="{gy}" x2="{W-pad}" y2="{gy}" stroke="#e5e7eb" stroke-width="1"/>'

    pts = " ".join(f"{x},{y}" for x,y in zip(xs,ys))
    poly = f'<polyline points="{pts}" fill="none" stroke="#2563eb" stroke-width="3"/>'
    dots = "".join(f'<circle cx="{x}" cy="{y}" r="5" fill="#2563eb"/>' for x,y in zip(xs,ys))

    val_labels = "".join(
        f'<text x="{x}" y="{y-6}" text-anchor="middle" font-size="{point_fs+4}" font-weight="700" fill="#0f172a">{int(v)}</text>'
        for x,y,v in zip(xs,ys,vals)
    )

    # نمایش تاریخ‌ها در دو خط: "X [ماه] تا" و "Y [ماه]" با راست‌چین
    week_labels = ""
    for x, lbl in zip(xs, weeks):
        # بررسی فرمت جدید با START: و END:
        if "START:" in lbl and "END:" in lbl:
            parts = lbl.split("|")
            start_part = parts[0].replace("START:", "").strip()
            end_part = parts[1].replace("END:", "").strip()
            # خط اول: تاریخ شروع با "تا" - راست‌چین
            week_labels += f'<text x="{x}" y="{H-22}" text-anchor="middle" direction="rtl" font-size="{max(10, point_fs-2)}" font-weight="600" fill="#475569">{start_part} تا</text>'
            # خط دوم: تاریخ پایان - راست‌چین
            week_labels += f'<text x="{x}" y="{H-6}" text-anchor="middle" direction="rtl" font-size="{max(10, point_fs-2)}" font-weight="600" fill="#475569">{end_part}</text>'
        else:
            # fallback برای فرمت قدیمی
            week_labels += f'<text x="{x}" y="{H-10}" text-anchor="middle" direction="rtl" font-size="{max(10, point_fs-2)}" font-weight="600" fill="#475569">{lbl}</text>'

    svg = f"""
<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}">
  {bg}
  {grid}
  {poly}
  {dots}
  {val_labels}
  {week_labels}
</svg>
"""
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode()).decode()

def build_acq_trend_svg_uri(df: pd.DataFrame, mp: dict, agent_id,
                            width: int = 260, height: int = 120,
                            point_fs: int = 12) -> str | None:
    """SVG کوچک خطی با برچسب بازه تاریخ و لیبل مقدار بالای هر نقطه."""
    try:
        # استفاده از همان تابع build_acq_trend_last4weeks برای دریافت داده‌ها با برچسب تاریخ
        agg = build_acq_trend_last4weeks(df, mp, agent_id)
        if agg is None or agg.empty:
            return None

        vals = agg["acq"].astype(float).tolist()
        disp_labels = agg["label"].tolist()

        W, H = width, height
        pad = max(24, point_fs * 2) 
        label_h = 35     # افزایش ارتفاع برای تاریخ‌های طولانی‌تر
        top_pad  = 18    # جا برای لیبلِ مقدار بالای نقطه
        plot_h = H - (pad + label_h + top_pad)
        plot_y0 = top_pad
        plot_y1 = top_pad + plot_h

        n = len(vals)
        if n == 1:
            xs = [W/2]
        else:
            step = (W - 2*pad) / (n - 1)
            xs = [pad + i*step for i in range(n)]

        vmin, vmax = min(vals), max(vals)
        if vmin == vmax:
            vmin -= 1.0; vmax += 1.0

        def y_map(v):
            t = (v - vmin) / (vmax - vmin) if vmax != vmin else 0.5
            return plot_y1 - t * plot_h

        ys = [y_map(v) for v in vals]

        bg = f'<rect x="0" y="0" width="{W}" height="{H}" fill="#ffffff"/>'

        # خطوط شبکه ملایم
        grid = ""
        for gy in [plot_y0, (plot_y0+plot_y1)/2, plot_y1]:
            grid += f'<line x1="{pad}" y1="{gy:.1f}" x2="{W-pad}" y2="{gy:.1f}" stroke="#e5e7eb" stroke-width="1"/>'

        pts = " ".join(f"{x:.1f},{y:.1f}" for x,y in zip(xs,ys))
        poly = f'<polyline points="{pts}" fill="none" stroke="#2563eb" stroke-width="3.5"/>'
        dots = "".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4.5" fill="#2563eb"/>' for x,y in zip(xs,ys))

        # لیبل مقدار بالای هر نقطه (عدد جذب همان هفته)
        val_labels = "".join(
            f'<text x="{x:.1f}" y="{(y-6):.1f}" text-anchor="middle" font-size="{point_fs+2}" font-weight="700" fill="#0f172a">{int(v)}</text>'
            for x,y,v in zip(xs,ys,vals)
        )

        # لیبل تاریخ‌ها در محور X با فرمت "X [ماه] تا" و "Y [ماه]" با راست‌چین
        week_labels = ""
        for x, lbl in zip(xs, disp_labels):
            # بررسی فرمت جدید با START: و END:
            if "START:" in lbl and "END:" in lbl:
                parts = lbl.split("|")
                start_part = parts[0].replace("START:", "").strip()
                end_part = parts[1].replace("END:", "").strip()
                # خط اول: تاریخ شروع با "تا" - راست‌چین
                week_labels += f'<text x="{x:.1f}" y="{H-20}" text-anchor="middle" direction="rtl" font-size="{max(9, point_fs-3)}" font-weight="600" fill="#475569">{start_part} تا</text>'
                # خط دوم: تاریخ پایان - راست‌چین
                week_labels += f'<text x="{x:.1f}" y="{H-4}" text-anchor="middle" direction="rtl" font-size="{max(9, point_fs-3)}" font-weight="600" fill="#475569">{end_part}</text>'
            else:
                # fallback برای فرمت قدیمی
                week_labels += f'<text x="{x:.1f}" y="{H-8}" text-anchor="middle" direction="rtl" font-size="{max(9, point_fs-3)}" font-weight="600" fill="#475569">{lbl}</text>'

        svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">
  {bg}
  {grid}
  <g>{poly}{dots}{val_labels}</g>
  {week_labels}
</svg>'''
        return "data:image/svg+xml;base64," + base64.b64encode(svg.encode("utf-8")).decode()
    except Exception:
        return None

def acq_last_k_weeks_sum(df: pd.DataFrame, mp: dict, agent_id, k: int = 4) -> int:
    try:
        sub = df[[mp["id"], mp["week"], mp["acq"]]].copy()
        sub = sub[sub[mp["id"]].astype(str).str.strip() == str(agent_id).strip()]
        if sub.empty:
            return 0
        
        sub["_week_num"] = to_numeric_clean(sub[mp["week"]])
        sub["_acq_num"]  = to_numeric_clean(sub[mp["acq"]])
        sub = sub.dropna(subset=["_week_num"])
        if sub.empty:
            return 0
        
        # جایگزین کردن NaN در acq با 0
        sub["_acq_num"] = sub["_acq_num"].fillna(0)
        
        agg = sub.groupby("_week_num")["_acq_num"].sum().reset_index()
        agg = agg.rename(columns={"_week_num": "week", "_acq_num": "acq"})
        agg = agg.sort_values("week")
        
        weeks_sorted = agg["week"].unique().tolist()
        if not weeks_sorted:
            return 0
        
        last_k = weeks_sorted[-k:]
        result = agg[agg["week"].isin(last_k)]["acq"].sum()
        return int(result) if not pd.isna(result) else 0
    except Exception:
        return 0


def acq_last_week_sum(df: pd.DataFrame, mp: dict, agent_id, default_k: int = 4) -> int:
    """
    مجموع جذب «هفته‌ی آخر» برای این بازاریاب.
    اگر کمتر از ۱ هفته دیتا باشد → 0
    """
    try:
        sub = df[[mp["id"], mp["week"], mp["acq"]]].copy()
        sub = sub[sub[mp["id"]].astype(str).str.strip() == str(agent_id).strip()]
        if sub.empty:
            return 0

        sub["_week_num"] = to_numeric_clean(sub[mp["week"]])
        sub["_acq_num"]  = to_numeric_clean(sub[mp["acq"]])
        sub = sub.dropna(subset=["_week_num"])
        if sub.empty:
            return 0

        # جایگزین کردن NaN در acq با 0
        sub["_acq_num"] = sub["_acq_num"].fillna(0)

        agg = sub.groupby("_week_num")["_acq_num"].sum().reset_index()
        agg = agg.rename(columns={"_week_num": "week", "_acq_num": "acq"})
        agg = agg.sort_values("week")
        
        weeks_sorted = agg["week"].unique().tolist()
        if not weeks_sorted:
            return 0
        
        last_week = weeks_sorted[-1]
        result = agg[agg["week"] == last_week]["acq"].sum()
        return int(result) if not pd.isna(result) else 0
    except Exception:
        return 0

# ---------- Supervisor daily / overview ----------
def compute_supervisor_daily(df: pd.DataFrame, mp: dict) -> pd.DataFrame:
    # یک کپی از df برای ستون‌های موقتی
    df = df.copy()

    # ---- QC ناظران میدانی در سطح ردیف خام ----
    pq_col = mp.get("presence_q_daily")
    bq_col = mp.get("banner_q_daily")

    if pq_col or bq_col:
        def _row_field_q(x):
            vals = []
            if pq_col and pq_col in x:
                v = _to_pct_0_100(x[pq_col])
                if v is not None:
                    vals.append(v)
            if bq_col and bq_col in x:
                v = _to_pct_0_100(x[bq_col])
                if v is not None:
                    vals.append(v)
            if not vals:
                return None
            return sum(vals) / len(vals)
        df["_field_q_day"] = df.apply(_row_field_q, axis=1)
    else:
        df["_field_q_day"] = None

    # ---- QC راننده‌ها در سطح ردیف خام (میانگین حضور/آموزش) ----
    pres_col = mp.get("presence_score")
    edu_col  = mp.get("edu_score")

    if pres_col or edu_col:
        def _row_driver_q(x):
            vals = []
            if pres_col and pres_col in x:
                v = _to_pct_0_100(x[pres_col])
                if v is not None:
                    vals.append(v)
            if edu_col and edu_col in x:
                v = _to_pct_0_100(x[edu_col])
                if v is not None:
                    vals.append(v)
            if not vals:
                return None
            return sum(vals) / len(vals)
        df["_driver_q_day"] = df.apply(_row_driver_q, axis=1)
    else:
        df["_driver_q_day"] = None

    # ---- تجمیع در سطح «بازاریاب-روز» ----
    per_id_day = df.groupby([mp["id"], "_date"]).agg(
        team_size_sum=(mp["people_col"], "sum"),       # افراد تیم (TeamSize)
        assigned_sum=(mp["assigned_col"], "sum"),      # تخصیص داده شده (actualTeamSize)
        absence_sum=(mp["absence"], "sum") if mp.get("absence") else ("_date","size"),
        incomplete_sum=(mp["incomplete"], "sum"),
        fake_sum=(mp["fake_sub"], "sum") if mp.get("fake_sub") else ("_date","size"),
        shift_delay_sum=(mp["shift_delay"], "sum"),
        false_check_sum=(mp["false_check"], "sum") if mp.get("false_check") else ("_date","size"),
        golden_sum=(mp["goldentime"], "sum"),
        hours_sum=(mp["total_hours"], "sum"),
        supervisor_first=(mp["supervisor"], "first"),
        city_first=(mp["city"], "first"),
        tier_first=(mp["tier"], "first"),
        date_first=(mp["date"], "first"),
        field_q_day_mean=("_field_q_day", "mean"),
        driver_q_day_mean=("_driver_q_day", "mean"),
    ).reset_index()

    # فقط گلدن‌تایم‌هایی که incomplete ندارند
    per_id_day["golden_clean"] = per_id_day.apply(
        lambda r: r["golden_sum"] if pd.notna(r["golden_sum"]) and (pd.isna(r["incomplete_sum"]) or r["incomplete_sum"] == 0) else 0,
        axis=1
    )

    # ---- تجمیع در سطح «سرپرست-شهر-روز» ----
    sup_day = per_id_day.groupby(
        ["supervisor_first","city_first","tier_first","_date","date_first"]
    ).agg(
        People=("team_size_sum","sum"),            # افراد تیم
        Assigned=("assigned_sum","sum"),           # تخصیص داده شده
        NoParticipation=("absence_sum","sum"),     # روزهای بدون مشارکت
        Incomplete=("incomplete_sum","sum"),
        InvalidDistance=("fake_sum","sum"),
        ShiftDelay=("shift_delay_sum","sum"),
        FalseCheck=("false_check_sum","sum"),
        GoldenTime=("golden_clean","sum"),
        TotalWorkHours=("hours_sum","sum"),
        field_q_pct=("field_q_day_mean","mean"),   # ⬅ QC ناظران میدانی (درصد)
        driver_q_pct=("driver_q_day_mean","mean"), # ⬅ QC راننده‌ها (درصد عددی)
    ).reset_index().rename(columns={
        "supervisor_first":"سرپرست",
        "city_first":"شهر",
        "tier_first":"Tier",
        "_date":"تاریخ",
        "date_first":"miladi date",
    })

    sup_day = sup_day.rename(columns={
        "People":"افراد تیم",
        "Assigned":"تخصیص داده شده",
        "NoParticipation":"روزهای بدون مشارکت",
        "Incomplete":"مشارکت‌های ناقص",
        "InvalidDistance":"مشارکت خارج از محدوده مجاز",
        "ShiftDelay":"مشارکت با تاخیر",
        "FalseCheck":"خوداظهاری اشتباه",
        "GoldenTime":"ساعات طلایی",
        "TotalWorkHours":"ساعات کاری کل",
    })

    # میانگین ساعت کاری
    sup_day["میانگین ساعات کاری"] = sup_day.apply(
        lambda r: (r["ساعات کاری کل"]/r["تخصیص داده شده"]) if r["تخصیص داده شده"] else None, axis=1
    )

    # ---- QC ناظران میدانی: به‌صورت درصد (۰–۱۰۰) ----
    sup_day["QC ناظران میدانی"] = sup_day["field_q_pct"].apply(
        lambda v: fmt_val(_to_pct_0_100(v) if v is not None else None, percent=True)
    )

    # ---- QC راننده‌ها: عالی / خوب / متوسط / ضعیف بر اساس درصد میانگین ----
    def _driver_q_label_from_pct(v):
        p = _to_pct_0_100(v)
        if p is None:
            return "-"
        if p <= 30:
            return "ضعیف"
        elif p <= 70:
            return "خوب"
        else:
            return "عالی"

    sup_day["QC راننده ها"] = sup_day["driver_q_pct"].apply(_driver_q_label_from_pct)

    # ستون‌های موقتی را لازم نداریم در خروجی نهایی
    sup_day = sup_day.drop(columns=["field_q_pct","driver_q_pct"], errors="ignore")

    sup_day["تاریخ"] = pd.to_datetime(sup_day["تاریخ"]).dt.strftime("%Y-%m-%d")
    return sup_day

# ---------------------- Helper functions for supervisor overview ----------------------
def _qc_driver_avg_to_label(avg_val) -> str:
    """
    Convert the numeric average of driver QC scores (1–4) to its qualitative Persian label.

    Parameters
    ----------
    avg_val : float or str or None
        The average value computed from mapping driver QC categories to numbers
        (ضعیف→1, متوسط→2, خوب→3, عالی→4). May be `None` or NaN if no data.

    Returns
    -------
    str
        The qualitative label:

        - If `avg_val` is None or NaN → "-".
        - If `avg_val` < 2.0 → "ضعیف" (weak).
        - If `avg_val` < 3.0 → "متوسط" (medium).
        - If `avg_val` < 3.5 → "خوب" (good).
        - Otherwise → "عالی" (excellent).

    This logic ensures that an average like 2.6 is considered "متوسط" as
    requested by the user.
    """
    try:
        # Check for pandas/numpy NaN or None
        if avg_val is None or (hasattr(pd, "isna") and pd.isna(avg_val)):
            return "-"
        v = float(avg_val)
    except Exception:
        return "-"
    if v < 2.0:
        return "ضعیف"
    elif v < 3.0:
        return "متوسط"
    elif v < 3.5:
        return "خوب"
    else:
        return "عالی"

def compute_supervisor_overview_total(sup_day: pd.DataFrame) -> pd.DataFrame:
    ov = (sup_day.groupby(["سرپرست","شهر"])
          .agg(
              Days=("تاریخ","nunique"),
              Num_NoPart=("روزهای بدون مشارکت","sum"),
              Num_Shift=("مشارکت با تاخیر","sum"),
              Num_Incomp=("مشارکت‌های ناقص","sum"),
              Num_Gold=("ساعات طلایی","sum"),
              Num_Hours=("ساعات کاری کل","sum"),
              Den_Assigned=("تخصیص داده شده","sum"),
          ).reset_index())

    d = ov["Den_Assigned"].replace(0, pd.NA)

    nopart_rate = (ov["Num_NoPart"] / d) * 100
    shift_rate  = (ov["Num_Shift"]  / d) * 100
    inc_rate    = (ov["Num_Incomp"] / d) * 100
    gold_rate   = (ov["Num_Gold"]   / d) * 100
    avg_hours   = (ov["Num_Hours"]  / d)

    out = pd.DataFrame({
        "میانگین ساعات کاری":  avg_hours.apply(lambda v: fmt_val(v, nd=2)),
        "مشارکت‌های ناقص":    inc_rate.apply(lambda v: fmt_val(v, percent=True)),
        "مشارکت با تاخیر":     shift_rate.apply(lambda v: fmt_val(v, percent=True)),
        "ساعات طلایی شیفت":    gold_rate.apply(lambda v: fmt_val(v, percent=True)),
        "روزهای بدون مشارکت": nopart_rate.apply(lambda v: fmt_val(v, percent=True)),
        "actual assigned":    ov["Den_Assigned"],  # بعداً در رندر حذف می‌شود (UI فقط)
        "سرپرست":             ov["سرپرست"],
        "شهر":                ov["شهر"],
        "Days":               ov["Days"],
    })
    return out

BASE_CSS = """
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;600;700&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#ffffff; --border:#e5e7eb; --head:#eef2ff; --text:#0f172a; --muted:#64748b;
    --tier1:#e0f2fe; --tier2:#dcfce7; --tier3:#fff7ed;
    --shadow:0 8px 24px rgba(15,23,42,.08);
    --grad: linear-gradient(135deg,#1e3a8a 0%, #0ea5e9 60%, #22d3ee 100%);
    --paperTop:#ff8c3a;
  }
  body{font-family:'Vazirmatn',sans-serif;background:#f6f8fb;}

  /* Paper container */
  .paper{direction:rtl; max-width:1280px; margin:24px auto; padding:24px 20px; background:#fff;
         border-radius:16px; box-shadow:var(--shadow); border-top:6px solid var(--paperTop);}

  /* Header block */
  .sup-wrap{direction:rtl;color:var(--text);max-width:100%;margin:0 auto;}
  .info-grid{display:grid;grid-template-columns:1fr;gap:8px;margin:0}
  .info-card{background:#f8fafc;border:1px solid var(--border);border-radius:12px;box-shadow:0 2px 8px rgba(2,6,23,.04);
             padding:8px 10px;text-align:center;font-weight:700;font-size:14px}
  .lbl{color:#475569;font-weight:600;margin-left:6px;font-size:12px}

  .header-line{position:relative;display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
  .header-logo{position:absolute;top:-2px;right:4px;height:34px;object-fit:contain;border-radius:8px}

  /* Section titles */
  .section-title{background:#eef2ff;color:#334155;border:1px solid #dbeafe;border-radius:12px;
                 padding:8px 12px;text-align:center;font-weight:700;margin:10px 0}

  /* Tables */
  .table-wrap{width:100%; overflow-x:hidden; border-radius:14px; background:var(--bg); box-shadow:0 2px 10px rgba(2,6,23,.06) }
  table.sup{width:100%; border-collapse:separate; border-spacing:0; table-layout:fixed}
  .sup th,.sup td{border:1px solid var(--border); padding:10px 12px; text-align:center; vertical-align:middle;
                  white-space:normal; word-break:break-word; hyphens:auto; font-size:12px; color:#0f172a; min-width:120px; max-width:200px}
  .sup thead th{background:var(--head); font-weight:700; position:sticky; top:0; z-index:2}
  .cap{font-size:12px; color:var(--muted); margin:10px 0; text-align:center}
  .rtl-cell{ direction: rtl; text-align: right; }

  /* Tier badges */
  .badge{display:inline-block; padding:2px 10px; border-radius:999px; font-weight:700; font-size:11px; border:1px solid currentColor}
  .badge.tier1{color:#0284c7;background:#e0f2fe}
  .badge.tier2{color:#059669;background:#dcfce7}
  .badge.tier3{color:#ea580c;background:#fff7ed}

  /* Agent Card – مینیمال، سفید، وسط‌چین */
  .cards-container{
    direction:rtl;
    color:var(--text);
    max-width:1100px;
    margin:16px auto;
  }
  .card{
    border-radius:18px;
    overflow:hidden;
    background:#ffffff;
    box-shadow:var(--shadow);
    border:1px solid #e2e8f0;
    padding:8px 16px 12px;
  }
  .row{
    display:flex;
    gap:12px;
    border-top:1px solid rgba(226,232,240,.8);
    padding:14px 0;
    justify-content:space-between;
  }
  .row:first-child{
    border-top:none;
    padding-top:6px;
  }
  /* هدر نارنجی ملایم */
  .row.hdr{
    background:#fff7ed;
    border-radius:14px;
    padding:12px 16px;
    margin-bottom:4px;
    border:1px solid #fed7aa;
  }
  .cell{
    flex:1;
    display:flex;
    flex-direction:column;
    justify-content:center;
    align-items:center;
    text-align:center;
    min-height:52px;
  }
  .cell.chart-cell{
    flex:2;              /* چارت پهن‌تر و سمت چپ */
    align-items:center;
  }

  .cell.chart-cell img{
    width:100%;
    max-width:420px;
    height:140px;
    object-fit:contain;
  }
  .label{
    font-weight:500;
    font-size:13px;
    color:#64748b;
    margin-bottom:2px;
  }
  .value{
    font-weight:800;
    font-size:16px;
    color:#0f172a;
  }
  .subtext{
    margin-top:4px;
    font-size:12px;
    color:#9ca3af;
  }
</style>
"""


def _tier_row_class(row) -> str:
    t = str(row.get("Tier","")).replace(" ", "").lower()
    if t in ("tier1","1"): return "tier1-row"
    if t in ("tier2","2"): return "tier2-row"
    if t in ("tier3","3"): return "tier3-row"
    return ""

def df_to_html_table(df: pd.DataFrame, caption: str) -> str:
    cols = df.columns.tolist()
    thead = "<thead><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr></thead>"
    rows_html = []
    for _, r in df.iterrows():
        rows_html.append("<tr>{}</tr>".format("".join(f"<td>{r[c]}</td>" for c in cols)))
    tbody = "<tbody>" + "".join(rows_html) + "</tbody>"
    return f'<div class="table-wrap"><table class="sup">{thead}{tbody}</table></div>'

def build_header_block(supervisor: str, city_display: str, date_from: str, date_to: str, logo_src: str | None = None) -> str:
    logo_html = f'<img src="{logo_src}" alt="logo" class="header-logo">' if logo_src else ""
    return f"""
<div class="header-line">
  <div class="info-grid" style="flex:1;">
    <div class="info-card"><span class="lbl">سرپرست:</span>{supervisor} — <span class="lbl">شهر:</span>{city_display} — <span class="lbl">تاریخ گزارش:</span>{date_from} تا {date_to}</div>
  </div>
  {logo_html}
</div>
"""

def build_agent_date_block_shamsi(date_from_j: str, date_to_j: str, logo_src: str | None = None) -> str:
    logo_html = f'<img src="{logo_src}" alt="logo" style="height:36px;object-fit:contain;border-radius:8px;">' if logo_src else ""
    return f"""
<div class="sup-wrap">
  <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;margin:4px 0 10px;">
    <div class="info-grid" style="flex:1;">
      <div class="info-card"><span class="lbl">تاریخ گزارش:</span>{date_from_j} تا {date_to_j}</div>
    </div>
    {f'<div style="padding:6px 10px;border:1px solid var(--border);border-radius:12px;background:#fff;box-shadow:var(--shadow)' + f'">{logo_html}</div>' if logo_src else ""}
  </div>
</div>
"""

def build_top10_html(city_label: str, top_df: pd.DataFrame, date_from_j: str, date_to_j: str, logo_src: str | None = None) -> str:
    df = top_df.copy().reset_index(drop=True)
    df.insert(0, "رتبه", range(1, len(df) + 1))
    table_html = df_to_html_table(df, "")
    logo_html = f'<img src="{logo_src}" alt="logo" class="header-logo">' if logo_src else ""
    header = f"""
      <div class=\"header-line\">
        <div class=\"info-grid\" style=\"flex:1;\">
          <div class=\"info-card\">
            <span class=\"lbl\">Top 10</span> – <span class=\"lbl\">شهر:</span>{city_label} —
            <span class=\"lbl\">تاریخ گزارش:</span>{date_from_j} تا {date_to_j}
          </div>
        </div>
        {logo_html}
      </div>
    """
    html = (
        '<!DOCTYPE html><html lang="fa"><head><meta charset="utf-8">' +
        BASE_CSS +
        '</head><body>' +
        '<div class="paper"><div class="sup-wrap">' +
        header +
        '<div class="section-title">10 نفر برتر</div>' +
        table_html +
        '</div></div>' +
        '</body></html>'
    )
    return html

def _format_daily_detailed(dy: pd.DataFrame) -> pd.DataFrame:
    disp = dy.drop(columns=["سرپرست","شهر","miladi date","ساعات کاری کل"], errors="ignore").copy()

    # Badge برای Tier
    if "Tier" in disp.columns:
        def _badge(v):
            s = str(v).strip().lower()
            if "1" in s: return '<span class="badge tier1">tier1</span>'
            if "2" in s: return '<span class="badge tier2">tier2</span>'
            if "3" in s: return '<span class="badge tier3">tier3</span>'
            return f'<span class="badge">{v}</span>'
        disp["Tier"] = disp["Tier"].apply(_badge)

    # ترتیب ستون‌ها (QC ها بعد از میانگین ساعات کاری نمایش داده می‌شوند)
    leading = [c for c in ["تاریخ","Tier","افراد تیم","تخصیص داده شده"] if c in disp.columns]
    others  = [c for c in disp.columns if c not in leading]
    disp = disp[leading + others]

    for c in disp.columns:
        if c in ["تاریخ","Tier"]:
            continue
        if c == "میانگین ساعات کاری":
            disp[c] = disp[c].apply(lambda v: fmt_val(v, nd=2))
        elif c == "QC ناظران میدانی":
            # به‌صورت درصد
            disp[c] = disp[c].apply(lambda v: fmt_val(_to_pct_0_100(v), percent=True) if v not in [None, ""] else "-")
        elif c == "QC راننده ها":
            # لیبل متنی (عالی/خوب/متوسط/ضعیف) رو دست نمی‌زنیم
            continue
        else:
            disp[c] = disp[c].apply(lambda v: fmt_val(v, nd=2))
    return disp

def build_supervisor_html_persian(supervisor: str, city_display: str, date_from: str, date_to: str,
                                  overview_df: pd.DataFrame, daily_df: pd.DataFrame, logo_src: str | None = None) -> str:
    ov = overview_df.copy()
    dy = daily_df.copy()
    if "تاریخ" in dy.columns:
        dy["تاریخ"] = pd.to_datetime(dy["تاریخ"], errors="coerce").dt.date.apply(
            lambda d: to_jalali_words(d, persian_digits=True) if pd.notna(d) else "-"
        )
    if not ov.empty:
        ov = ov.drop(columns=["سرپرست","شهر","Days","actual assigned"], errors="ignore")

        cols = [
            "میانگین ساعات کاری",
            "مشارکت‌های ناقص",
            "مشارکت با تاخیر",
            "ساعات طلایی شیفت",
            "روزهای بدون مشارکت",
        ]

        # 🟦 اگر میانگین Total QC بود → اضافه کن
        if "میانگین TOTAL QC SCORE" in ov.columns:
            cols.append("میانگین TOTAL QC SCORE")

        # 🟩 میانگین QC ناظران میدانی
        if "میانگین QC ناظران میدانی" in ov.columns:
            cols.append("میانگین QC ناظران میدانی")

        # 🟨 میانگین QC راننده ها
        if "میانگین QC راننده ها" in ov.columns:
            cols.append("میانگین QC راننده ها")

        ov = ov[[c for c in cols if c in ov.columns]]

        # فرمت‌دهی
        for c in ov.columns:
            if c in ["میانگین ساعات کاری"]:
                ov[c] = ov[c].apply(lambda v: fmt_val(v, nd=2))
            elif c in ["میانگین QC ناظران میدانی"]:
                ov[c] = ov[c].apply(lambda v: fmt_val(v, percent=True))
            elif c in ["میانگین QC راننده ها"]:
                # به‌جای نمایش عدد، میانگین QC راننده‌ها را به برچسب کیفی تبدیل کن
                ov[c] = ov[c].apply(lambda v: _qc_driver_avg_to_label(v))
            elif c == "میانگین TOTAL QC SCORE":
                ov[c] = ov[c].apply(lambda v: fmt_val(v, nd=1))

        ov = add_badges_overview(ov)
    # این دو ستون فقط برای نمای کلی هفته استفاده شدند؛
    # در جدول جزئیات روزانه نمی‌خواهیم نمایش داده شوند.
    dy = dy.drop(columns=["QC ناظران میدانی", "QC راننده ها"], errors="ignore")
    dy = _format_daily_detailed(dy)
    dy = _format_daily_detailed(dy)    # حذف ستون‌های QC راننده و ناظر
    dy = dy.drop(columns=["نمره QC راننده ها","نمره QC ناظران میدانی"], errors="ignore")


    # PAPER + SECTIONS
    header_html = build_header_block(supervisor, city_display, date_from, date_to, logo_src=logo_src)
    html = BASE_CSS + '<div class="paper"><div class="sup-wrap">'
    html += header_html
    html += '<div class="section-title">نمای کلی هفته</div>'
    html += df_to_html_table(ov, "")
    html += '<div class="section-title">جزئیات روزانه (۷ روز اخیر)</div>'
    cols = dy.columns.tolist()
    thead = "<thead><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr></thead>"
    rows = []
    for _, r in dy.iterrows():
        tds = "".join(f"<td>{r[c]}</td>" for c in cols)
        rows.append(f"<tr>{tds}</tr>")
    day_html = f'<div class="table-wrap"><table class="sup">{thead}<tbody>{"".join(rows)}</tbody></table></div>'
    html += day_html
    html += '</div></div>'
    return html

def build_card_html(
    p: dict,
    date_from_j: str | None = None,
    date_to_j: str | None = None,
    logo_src: str | None = None,
    rank_text: str | None = None,
    team_rank_text: str | None = None,
    acq_chart_src: str | None = None,
    acq_recent3_total: int | None = None,
    acq_last_week: int | None = None,   # 👈 اضافه شد
) -> str:

    # --- داده‌های اصلی ---
    top_city = p.get("شهر","")
    top_sup  = p.get("سرپرست","")
    top_name = p.get("نام و نام خانوادگی","")

    rank_city  = rank_text or "-"
    rank_team  = team_rank_text or "-"

    active_drivers_count = fmt_val(p.get("تعداد راننده‌های فعال"))
    active_drivers_pct = fmt_val(p.get("درصد راننده‌های فعال"), percent=True)

    def _compact_jdate_range(df: str, dt: str) -> str:
        """
        ورودی ممکن است هر کدام از این ساختارها باشد:
          - '۱۴۰۴ آبان ۱۴'
          - '14 آبان 1404'
          - '14 آبان'
          - '1404 آبان 14'
          - 'آبان 14 1404'
          - و حتی اشتباه‌چین مثل: '1404 آبان' + '14 آبان 1404'

        خروجی استاندارد:
          '14 تا 20 آبان 1404'
        """
        import re

        # تبدیل فارسی → انگلیسی
        trans = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
        df = df.translate(trans)
        dt = dt.translate(trans)

        # استخراج عددها و ماه‌ها
        MONTHS = ["فروردین","اردیبهشت","خرداد","تیر","مرداد","شهریور",
                  "مهر","آبان","آذر","دی","بهمن","اسفند"]

        def parse(jstr):
            """تشخیص خودکار روز/ماه/سال از یک رشتهٔ فارسی"""
            parts = jstr.strip().split()

            day = None
            month = None
            year = None

            for p in parts:
                if p.isdigit():
                    # روز یا سال؟
                    if len(p) == 4:  # سال
                        year = p
                    else:
                        day = p
                elif p in MONTHS:
                    month = p

            return day, month, year

        d1,m1,y1 = parse(df)
        d2,m2,y2 = parse(dt)

        # اگر ماه/سال در ابتدا نبود اما در دومی بود → از دومی بردار
        if not m1: m1 = m2
        if not y1: y1 = y2
        if not y2: y2 = y1

        # اگر باز هم چیزی خالی بود → خروجی ساده
        if not (d1 and d2 and m1 and m2 and y1):
            return f"{df} تا {dt}"

        # اگر ماه‌ها متفاوت هستند، هر دو ماه را نمایش بده
        if m1 != m2:
            return f"{d1} {m1} تا {d2} {m2} {y1}"
        else:
            return f"{d1} تا {d2} {m1} {y1}"


    date_str = _compact_jdate_range(date_from_j, date_to_j) if (date_from_j and date_to_j) else ""


    no_part_days   = fmt_val(p.get("روزهای بدون مشارکت"))
    part_partial   = fmt_val(p.get("مشارکت‌های ناقص"))
    invalid_dist   = fmt_val(p.get("مشارکت خارج از محدوده مجاز"))
    shift_delay    = fmt_val(p.get("مشارکت با تاخیر"))
    false_check    = fmt_val(p.get("خوداظهاری اشتباه"))

    # داده‌های نقص مدارک و راننده‌های فعال
    doc_issue_total = fmt_val(p.get("تعداد نقص مدارک"))
    doc_issue_pct   = fmt_val(p.get("درصد نقص مدارک"), percent=True)
    active_drivers_count = fmt_val(p.get("تعداد راننده‌های فعال"))
    active_drivers_pct = fmt_val(p.get("درصد راننده‌های فعال"), percent=True)

    # بولت‌پوینت جزئیات نقص مدارک (بر اساس ستون‌های نقص مدارک)
    doc_issue_items = []
    # اگر تعداد نقص مدارک صفر باشد، هیچ بولت‌پوینتی نمایش نده
    show_detail = True
    try:
        raw_doc_count = p.get("تعداد نقص مدارک")
        if raw_doc_count is not None:
            cval = float(raw_doc_count)
            if cval == 0:
                show_detail = False
    except Exception:
        show_detail = True
    # در صورت مجاز بودن، آیتم‌ها را بررسی و اضافه کن
    if show_detail:
        for col in DOC_ISSUE_COLS:
            if col in p:
                raw_val = p.get(col)
                try:
                    v = float(str(raw_val).replace(",", ""))
                except Exception:
                    v = None
                if v is not None and v > 0:
                    v_int = int(v)
                    v_str = str(v_int) if abs(v - v_int) < 1e-9 else f"{v:.0f}"
                    doc_issue_items.append(f"{col} – {v_str}")

    if doc_issue_items:
        doc_issues_html = (
            "<ul style='margin:6px 0 0; padding:0; color:var(--muted); "
            "font-size:12px; list-style:disc inside; direction:rtl; text-align:right;'>"
            + "".join(
                f"<li style='list-style-position:inside;'>{txt}</li>"
                for txt in doc_issue_items
            )
            + "</ul>"
        )
    else:
        doc_issues_html = ""

    acq_recent4    = fmt_val(acq_recent3_total) if acq_recent3_total is not None else "-"
    last_week_val  = fmt_val(acq_last_week) if acq_last_week is not None else "-"

    qc_driver_lbl  = p.get("نمره QC راننده ها", "ok")
    qc_field_lbl   = fmt_val(p.get("نمره QC ناظران میدانی"), percent=True)

    # --- helper برای تبدیل به ۰–۱ ---
    def _to_01(v):
        try:
            if v is None:
                return None
            import math
            if isinstance(v, float) and math.isnan(v):
                return None
            s = str(v).strip().replace(",", "")
            if s.endswith("%"):
                s = s[:-1].strip()
            if s == "" or s.lower() == "nan":
                return None
            x = float(s)
            if x > 1.01:
                x = x / 100.0
            return max(0.0, min(1.0, x))
        except Exception:
            return None

    # --- نمره‌های کلی حضور/آموزش (۰–۱۰۰) که قبلاً داشتیم ---
    def _to_num100(v):
        import math
        try:
            if v is None: return None
            if isinstance(v, float) and math.isnan(v): return None
            s = str(v).strip().replace(",", "")
            if s.endswith("%"): s = s[:-1].strip()
            if s == "" or s.lower() == "nan": return None
            x = float(s)
            if x <= 1.01: x *= 100.0
            return x
        except Exception:
            return None

    edu_v  = _to_num100(p.get("نمره آموزش توسط راننده ها"))
    pres_v = _to_num100(p.get("نمره حضور توسط راننده ها"))
    field_v= _to_num100(p.get("نمره QC ناظران میدانی"))

    # آستانه‌ها
    TH_EDU = 70.0
    TH_PRES = 70.0

    # ستون‌های جزئی مرتبط با حضور
    presence_detail_cols = {
        "visit": "مراجعه",
        "presence_status": "وضعیت حضور",
        "appearance": "پوشش و رفتار مناسب",
        "signup_time_p": "مدت زمان ثبت نام",
    }
    # ستون‌های جزئی مرتبط با آموزش
    edu_detail_cols = {
        "signup_time_e": "مدت زمان ثبت نام",
        "benefits": "مزایا",
        "daily_settlement": "تسویه روزانه",
        "support": "پشتیبانی",
        "unreal": "وعده ی غیرواقعی",
    }

    def _pick_weakest(detail_map):
        """کمترین نمره بین ستون‌های جزئی (۰–۱)."""
        weakest_key = None
        weakest_val = None
        for key, col in detail_map.items():
            v = _to_01(p.get(col))
            if v is None:
                continue
            if weakest_val is None or v < weakest_val:
                weakest_val = v
                weakest_key = key
        return weakest_key, weakest_val

    # پیغام‌های پیشنهادی برای هر آیتم
    def _presence_msg_for(key):
        if key == "visit":
            return "بهتره فعال‌تر به سراغ راننده‌ها بری و تعداد مراجعات مستقیمت رو بیشتر کنی.🤗💡"
        if key == "presence_status":
            return "بهتره بیشتر از ماشین بیرون باشی تا بتونی بیشتر هم جذب کنی.🧍‍♂️🧍‍♀️"
        if key == "appearance":
            return "روی پوشش و رفتار حرفه‌ای‌تر کار کن؛ ظاهر و برخورد خوب، اعتماد راننده رو زیاد می‌کنه.👔✨"
        if key == "signup_time_p":
            return "سعی کن زمان ثبت‌نام راننده رو کوتاه‌تر و روان‌تر کنی تا وسط کار منصرف نشه.😉👍"
        return "حضور میدانی‌ات رو منظم‌تر و فعال‌تر کن تا جذب بیشتری بگیری.🚶‍♂️✨"

    def _edu_msg_for(key):
        if key == "signup_time_e":
            return "مراحل ثبت‌نام رو مرحله‌به‌مرحله و واضح‌تر برای راننده توضیح بده.🚗💬"
        if key == "benefits":
            return "مزایای همکاری با ما رو شفاف‌تر و جذاب‌تر برای راننده توضیح بده.😎👌"
        if key == "daily_settlement":
            return "شرایط و زمان تسویه رو با مثال‌های ساده توضیح بده تا هیچ ابهامی نمونه.📆💵"
        if key == "support":
            return "بیشتر روی معرفی تیم پشتیبانی و راه‌های کمک‌گرفتن راننده‌ها تأکید کن.🚗💬"
        if key == "unreal":
            return "از وعده‌های خیلی خوش‌بینانه یا غیرواقعی پرهیز کن؛ اعتماد راننده مهم‌تر از جذب سریع است.🤝🙂"
        return "روی توضیح مزایا و شرایط همکاری بیشتر کار کن تا راننده قانع‌تر بشه.🔄💬"

    driver_qc_notes: list[str] = []

    # --- منطق حضور (Presence) ---
    if pres_v is not None:
        weakest_p_key, weakest_p_val = _pick_weakest(presence_detail_cols)
        if pres_v < TH_PRES and weakest_p_key is not None:
            driver_qc_notes.append(_presence_msg_for(weakest_p_key))
        else:
            driver_qc_notes.append("حضور میدانی‌ات خوب بوده؛ همین روند رو ادامه بده 🚀✨")

    # --- منطق آموزش (Education) ---
    if edu_v is not None:
        weakest_e_key, weakest_e_val = _pick_weakest(edu_detail_cols)
        if edu_v < TH_EDU and weakest_e_key is not None:
            driver_qc_notes.append(_edu_msg_for(weakest_e_key))
        else:
            driver_qc_notes.append("نحوهٔ توضیح شرایط و مزایا به راننده‌ها خوب بوده؛ همین سبک رو حفظ کن 👏🌟")

    driver_qc_notes = driver_qc_notes[:2]
    if not driver_qc_notes:
        driver_qc_notes = ["مزایای همکاری با ما رو شفاف‌تر و جذاب‌تر برای راننده توضیح بده.🤝🙂"]

    # --- QC ناظران میدانی: بولت‌پوینت‌های اختصاصی ---
    field_qc_notes = []

    # استخراج درصدها از 0–1 یا 0–100
    pres_q_pct = _to_num100(p.get("presenceQuality_daily"))
    bann_q_pct = _to_num100(p.get("bannerQuality_daily"))

    # presenceQuality_daily → کیفیت حضور میدانی
    if pres_q_pct is not None:
        if pres_q_pct < 70:
            field_qc_notes.append(
                "بهتره زمان بیشتری خارج از خودرو باشی تا حضور میدانی‌ات بیشتر دیده بشه.🚶‍♂️✨"
            )
        elif pres_q_pct >= 90:
            field_qc_notes.append(
                "حضور میدانی‌ات عالی بوده؛ همین روند رو ادامه بده.👏🌟"
            )

    # bannerQuality_daily → کیفیت نصب و دید بنر
    if bann_q_pct is not None:
        if bann_q_pct < 90:
            field_qc_notes.append(
                "لازمه بنرت رو طوری نصب کنی که در دید راننده‌ها باشه و از نظر کیفیت واضح دیده بشه.📣👌"
            )
        elif bann_q_pct >= 95:
            field_qc_notes.append(
                "کیفیت نصب بنرت عالی بوده؛ همین استاندارد رو حفظ کن.📣🌟"
            )

    # اگر هیچ پیام مستقیمی نبود ولی نمره QC ناظر داریم
    if not field_qc_notes and field_v is not None:
        if float(field_v) >= 80:
            field_qc_notes.append("حضور و اجرای میدانی‌ات خوب بوده؛ همین روند را حفظ کن.✅")
        else:
            field_qc_notes.append("حضور و همراهی میدانی‌ات می‌تونه فعال‌تر باشه تا نتیجهٔ بهتری بگیری.🤝")

    logo_html = f'<img src="{logo_src}" alt="logo" style="height:40px;object-fit:contain;border-radius:8px;">' if logo_src else ""
    total_qc_score_val = fmt_val(p.get("TOTAL QC SCORE"), nd=1)

    chart_html = f'<img src="{acq_chart_src}" alt="acq trend">' if acq_chart_src else ""

    # ---- HTML بولت‌پوینت‌ها ----
    notes_html = "".join(
        f"<li style='list-style-position:inside;'>{txt}</li>" for txt in driver_qc_notes
    )
    driver_qc_ul = f"<ul style='margin:6px 0 0; padding:0; color:var(--muted); font-size:12px; list-style:disc inside; direction:rtl; text-align:right;'>{notes_html}</ul>"

    # ساخت HTML بولت‌ها
    if field_qc_notes:
        notes_html_field = "".join(
            f"<li style='list-style-position:inside;'>{txt}</li>"
            for txt in field_qc_notes
        )
        field_qc_ul = (
            "<ul style='margin:6px 0 0; padding:0; color:var(--muted); "
            "font-size:12px; list-style:disc inside; direction:rtl; text-align:right;'>"
            f"{notes_html_field}</ul>"
        )
    else:
        field_qc_ul = ""

    # ---- ساخت کارت ----
    return f"""
{BASE_CSS}
<div class="cards-container">
  <div class="card">

    <!-- سطر ۱: هدر -->
    <div class="row hdr">
      <div class="cell">
        <div class="label">نام</div>
        <div class="value">{top_name}</div>
        {f'<div class="subtext">{date_str}</div>' if date_str else ""}
      </div>
      <div class="cell">
        <div class="label">شهر</div>
        <div class="value">{top_city}</div>
      </div>
      <div class="cell">
        <div class="label">سرپرست</div>
        <div class="value">{top_sup}</div>
      </div>
      <div class="cell">
        <div class="label">رتبه در تیم</div>
        <div class="value">{rank_team}</div>
      </div>
      <div class="cell">
        <div class="label">رتبه در شهر</div>
        <div class="value">{rank_city}</div>
      </div>
      <div class="cell" style="flex:0 0 90px;">
        {logo_html}
      </div>
    </div>

    <!-- سطر ۲: متریک‌ها (بدون اکتیو و تطبیق) -->
    <div class="row">
      <div class="cell">
        <div class="label">روزهای بدون مشارکت</div>
        <div class="value">{no_part_days}</div>
      </div>
      <div class="cell">
        <div class="label">مشارکت‌های ناقص</div>
        <div class="value">{part_partial}</div>
      </div>
      <div class="cell">
        <div class="label">مشارکت خارج از محدوده مجاز</div>
        <div class="value">{invalid_dist}</div>
      </div>
      <div class="cell">
        <div class="label">مشارکت با تاخیر</div>
        <div class="value">{shift_delay}</div>
      </div>
      <div class="cell">
        <div class="label">خوداظهاری اشتباه</div>
        <div class="value">{false_check}</div>
      </div>
    </div>

    <!-- سطر ۳: QC -->
    <div class="row">
      <div class="cell">
        <div class="label">نمره QC راننده ها</div>
        <div class="value">{qc_driver_lbl}</div>
        {driver_qc_ul}
      </div>
      <div class="cell">
        <div class="label">نمره QC ناظران میدانی</div>
        <div class="value">{qc_field_lbl}</div>
        {field_qc_ul}
      </div>
      <div class="cell">
        <div class="label">TOTAL QC SCORE</div>
        <div class="value">{total_qc_score_val}</div>
      </div>
    </div>

    <!-- سطر ۴: سمت راست متن‌ها، سمت چپ ترند ۴ هفته‌ای -->
    <div class="row">
      <!-- سمت راست (نقص مدارک و راننده‌های فعال) -->
      <div class="cell" style="align-items:flex-start;text-align:right;">
        <div class="label">تعداد نقص مدارک</div>
        <div class="value">{doc_issue_total}</div>
        {doc_issues_html}

        <div class="label" style="margin-top:8px;">درصد نقص مدارک</div>
        <div class="value">{doc_issue_pct}</div>

        <div class="label" style="margin-top:12px;">تعداد راننده‌های فعال</div>
        <div class="value">{active_drivers_count}</div>

        <div class="label" style="margin-top:8px;">درصد راننده‌های فعال</div>
        <div class="value">{active_drivers_pct}</div>
      </div>

      <!-- سمت چپ (نمودار ترند ۴ هفته اخیر) -->
      <div class="cell chart-cell">
        <div class="label" style="margin-bottom:4px;">
          روند جذب ۴ هفته اخیر
        </div>
        <div class="label" style="margin-bottom:6px;font-weight:600;color:#0f172a;">
          مجموع جذب هفته پیش:
          <span style="margin-right:4px;font-weight:800;color:#0f172a;">{last_week_val}</span>
        </div>
        {chart_html}
      </div>
    </div>
  </div>
</div>
"""




def html_download_button(filename: str, html_str: str, label: str):
    b64 = base64.b64encode(html_str.encode("utf-8")).decode()
    st.markdown(f'<a download="{filename}" href="data:text/html;base64,{b64}">{label}</a>', unsafe_allow_html=True)

def image_to_b64(file) -> str | None:
    if not file:
        return None
    ext = (file.name.split(".")[-1] or "").lower()
    mime = "image/png" if ext in ["png"] else "image/jpeg" if ext in ["jpg","jpeg"] else "image/svg+xml" if ext=="svg" else "image/*"
    data = file.read()
    b64 = base64.b64encode(data).decode()
    return f"data:{mime};base64,{b64}"

# ===== Top performer scoring (پایه) =====
_QC_DRIVER_SCORE = {
    "عالی": 1.00,
    "خوب": 0.60,
    "متوسط": 0.50,
    "ضعیف": 0.20,
}
def _percent_to_0_1(v):
    if v is None or (hasattr(pd,"isna") and pd.isna(v)): return float('nan')
    s = str(v).strip().replace(",","")
    if s.endswith("%"): s = s[:-1].strip()
    if s == "": return float('nan')
    try:
        x = float(s)
        if x <= 1.01: return max(0.0, min(1.0, x))
        return max(0.0, min(1.0, x/100.0))
    except Exception:
        return float('nan')

def _minmax(series: pd.Series):
    s = pd.to_numeric(series, errors="coerce")
    mn, mx = s.min(), s.max()
    if pd.isna(mn) or pd.isna(mx) or mx - mn == 0:
        return pd.Series([0.0]*len(s), index=series.index)
    return (s - mn) / (mx - mn)

import math  # اگر بالاتر import نشده


def compute_agent_rank_score(stats: pd.DataFrame) -> pd.DataFrame:
    """
    ساخت امتیاز رتبه برای هر بازاریاب بر اساس منطق اصلاح‌شده:

    - ۸۰٪ → مؤلفه جذب (APA): تعداد جذب تقسیم بر روزهای کاری، سپس نرمال‌سازی min–max.
      این مؤلفه عملکرد جذب را در دورهٔ موردنظر (معمولاً بازهٔ اخیر) می‌سنجد.

    - ۲۰٪ → مؤلفه کیفیت عملکرد بر اساس TOTAL QC SCORE:
      ابتدا TOTAL QC SCORE که عددی بین ۰ تا ۱۰۰ است به یک نمرهٔ گسسته در بازهٔ ۰ تا ۲۰ تبدیل می‌شود:
        * اگر TOTAL QC SCORE ≥ 80: نمرهٔ 20
        * اگر 60 ≤ TOTAL QC SCORE < 80: نمرهٔ 10
        * اگر 50 ≤ TOTAL QC SCORE < 60: نمرهٔ 5
        * در غیر این صورت: نمرهٔ 0

      سپس این نمرهٔ 0–20 بر 20 تقسیم می‌شود تا به بازهٔ ۰–۱ نگاشت شود. امتیاز نهایی با وزن ۰٫۲
      ترکیب می‌شود.

    نتیجهٔ نهایی (Score) برابر است با ۰٫۸ × (نمره نرمال‌شدهٔ جذب) + ۰٫۲ × (نمره نرمال‌شدهٔ کیفیت).
    """
    df = stats.copy()
    if df.empty:
        df["Score"] = 0.0
        df["_ACQ_rank"] = 0.0
        df["_QC_component"] = 0.0
        return df

    # --- ۸۰٪: جذب به ازای روز کاری ---
    # محاسبه APA و نرمال‌سازی آن به بازهٔ ۰–۱
    acq = pd.to_numeric(df.get("تعداد جذب"), errors="coerce")
    wd  = pd.to_numeric(df.get("روزهای کاری"), errors="coerce")
    # جلوگیری از تقسیم بر صفر
    acq_per_day = acq / wd.replace(0, pd.NA)
    df["_ACQ_rank"] = _minmax(acq_per_day).fillna(0.0)

    # --- ۲۰٪: مؤلفه کیفیت بر اساس TOTAL QC SCORE ---
    # نرمال‌سازی TOTAL QC SCORE بر اساس محدوده‌های گسسته
    qc_component = pd.Series(0.0, index=df.index)
    if "TOTAL QC SCORE" in df.columns:
        tot_qc = pd.to_numeric(df["TOTAL QC SCORE"], errors="coerce")
        def qc_map(x: float | int | None) -> float:
            if pd.isna(x):
                return 0.0
            try:
                val = float(x)
            except Exception:
                return 0.0
            # تبدیل بر اساس محدوده‌ها
            if val >= 80:
                return 1.0  # 20/20
            elif val >= 60:
                return 0.5  # 10/20
            elif val >= 50:
                return 0.25  # 5/20
            else:
                return 0.0
        qc_component = tot_qc.apply(qc_map)
    df["_QC_component"] = qc_component

    # ترکیب نهایی: ۸۰٪ جذب، ۲۰٪ کیفیت
    df["Score"] = 0.8 * df["_ACQ_rank"] + 0.2 * df["_QC_component"]
    return df

def _to_pct_0_100(v):
    """
    ورودی می‌تونه:
      - ۰–۱ (مثلاً 0.8)
      - یا ۰–۱۰۰ (مثلاً 80 یا '80%')
    خروجی همیشه درصد (۰ تا ۱۰۰) است.
    """
    if v is None or (hasattr(pd, "isna") and pd.isna(v)):
        return None
    s = str(v).strip().replace(",", "")
    if s.endswith("%") or s.endswith("٪"):
        s = s[:-1].strip()
    if s == "" or s.lower() == "nan":
        return None
    try:
        x = float(s)
        if x <= 1.01:   # فرض: ۰–۱ بوده
            x *= 100.0
        return x
    except Exception:
        return None

# ===== TOTAL QC SCORE (rule-based) =====

def score_false_check(pct_false: float) -> float:
    """درصد خوداظهاری اشتباه (۰ تا ۱۰۰) → امتیاز بین ۰ تا ۲۰ (برعکس)"""
    if pct_false is None or (hasattr(pd, "isna") and pd.isna(pct_false)):
        return 0.0
    good = 100.0 - float(pct_false)  # هرچه کمتر خطا، good بیشتر
    if good < 70:
        return 0.0
    elif good < 80:
        return 5.0
    elif good < 90:
        return 10.0
    elif good < 99:
        return 15.0
    else:
        return 20.0


def score_goldentime(zero_golden_count: float) -> float:
    """
    تعداد دفعاتی که goldentime = 0 بوده در بازه → امتیاز ۰ تا ۷
      0 بار صفر → 7 امتیاز
      1 یا 2 بار صفر → 5 امتیاز
      3 بار صفر → 3 امتیاز
      4+ بار صفر → 0 امتیاز
    """
    if zero_golden_count is None or (hasattr(pd,"isna") and pd.isna(zero_golden_count)):
        return 0.0

    z = float(zero_golden_count)

    if z == 0:
        return 7.0
    elif z <= 2:
        return 5.0
    elif z == 3:
        return 3.0
    else:
        return 0.0



def score_total_work_hours(avg_hours: float) -> float:
    """میانگین ساعات کاری در روز → امتیاز ۰ تا ۸"""
    if avg_hours is None or (hasattr(pd,"isna") and pd.isna(avg_hours)):
        return 0.0

    h = float(avg_hours)

    if h > 8:
        return 8.0
    elif h >= 7.5:
        return 6.0
    elif h >= 7.0:   # ← حالا 7 هم شامل میشه
        return 4.0
    else:
        return 0.0   # زیر 7 → صفر


def score_shift_delay(n_shift_delay: float) -> float:
    """تعداد مشارکت با تاخیر در بازه → امتیاز ۰ تا ۸"""
    if n_shift_delay is None or (hasattr(pd,"isna") and pd.isna(n_shift_delay)):
        return 0.0
    d = float(n_shift_delay)
    if d == 0:
        return 8.0
    elif d <= 1.5:
        return 6.0
    elif d <= 3.0:
        return 4.0
    else:
        return 0.0

def score_incomplete(n_incomplete: float) -> float:
    """تعداد مشارکت ناقص در بازه → امتیاز ۰ تا ۲"""
    if n_incomplete is None or (hasattr(pd,"isna") and pd.isna(n_incomplete)):
        return 0.0
    n = float(n_incomplete)
    if n == 0:
        return 2.0
    elif n <= 3:
        return 1.0
    else:
        return 0.0

def score_banner_daily(pct: float) -> float:
    """
    banner_daily (درصد) → حداکثر ۵ امتیاز
    >40% → 5 امتیاز
    20–40% → 2.5 امتیاز
    بقیه → 0
    """
    if pct is None or (hasattr(pd,"isna") and pd.isna(pct)):
        return 0.0
    p = float(pct)
    if p >= 40:
        return 5.0
    elif p >= 20:
        return 2.5
    else:
        return 0.0

def score_banner_quality(pct: float) -> float:
    """
    bannerQuality_daily (درصد) → حداکثر ۵ امتیاز
    >90% → 5
    70–90% → 2
    بقیه → 0
    """
    if pct is None or (hasattr(pd,"isna") and pd.isna(pct)):
        return 0.0
    p = float(pct)
    if p > 90:
        return 5.0
    elif p >= 70:
        return 2.0
    else:
        return 0.0

def score_presence_quality(pct: float) -> float:
    """
    presenceQuality_daily (درصد) → حداکثر ۵ امتیاز
    >70% → 5
    50–70% → 2
    بقیه → 0q
    """
    if pct is None or (hasattr(pd,"isna") and pd.isna(pct)):
        return 0.0
    p = float(pct)
    if p > 70:
        return 5.0
    elif p >= 50:
        return 2.0
    else:
        return 0.0

def score_presence_score(pct: float) -> float:
    """
    Presence Score راننده‌ها (درصد) → حداکثر ۵ امتیاز
    >70% → 5
    50–70% → 2
    بقیه → 0
    """
    if pct is None or (hasattr(pd,"isna") and pd.isna(pct)):
        return 0.0
    p = float(pct)
    if p > 70:
        return 5.0
    elif p >= 50:
        return 2.0
    else:
        return 0.0

def score_education_score(pct: float) -> float:
    """
    Eduction Score راننده‌ها (درصد) → حداکثر ۱۰ امتیاز
    >90% → 10
    80–90% → 8
    70–80% → 6
    بقیه → 0
    """
    if pct is None or (hasattr(pd,"isna") and pd.isna(pct)):
        return 0.0
    p = float(pct)
    if p > 90:
        return 10.0
    elif p >= 80:
        return 8.0
    elif p >= 70:
        return 6.0
    else:
        return 0.0



def score_tatbigh(pct: float) -> float:
    """
    Tatbigh (درصد نقص مدارک؛ هرچه کمتر بهتر) → حداکثر ۵ امتیاز
    0% → 5
    0–5% → 4
    5–10% → 3
    >10% → 0
    """
    if pct is None or (hasattr(pd,"isna") and pd.isna(pct)):
        return 0.0
    p = float(pct)
    if p <= 0:
        return 5.0
    elif p <= 5:
        return 4.0
    elif p <= 10:
        return 3.0
    else:
        return 0.0

def score_activation(pct: float) -> float:
    """
    Activation (درصد راننده‌های فعال) → حداکثر ۵ امتیاز
    >80% → 5
    60–80% → 2.5
    بقیه → 0
    """
    if pct is None or (hasattr(pd,"isna") and pd.isna(pct)):
        return 0.0
    p = float(pct)
    if p > 80:
        return 5.0
    elif p >= 60:
        return 2.5
    else:
        return 0.0

# ===== New absence scoring for TOTAL QC SCORE =====
def score_absence_total(absence: float) -> float:
    """
    Absence (number of days without participation) → امتیاز ۰ یا ۵.

    بر اساس منطق جدید:
      - اگر absence یا «روزهای بدون مشارکت» کاملاً صفر باشد → امتیاز کامل (۵ امتیاز)
      - در غیر این صورت → ۰ امتیاز

    ورودی absence می‌تواند None یا عددی باشد؛ اگر None باشد یا NaN به ۰ در نظر گرفته می‌شود.
    """
    # تبدیل مقدار به عدد
    if absence is None or (hasattr(pd, "isna") and pd.isna(absence)):
        # در صورت خالی بودن، absence = 0
        n = 0.0
    else:
        try:
            n = float(absence)
        except Exception:
            n = 0.0
    # اگر هیچ غیبتی وجود نداشته باشد (absense == 0) امتیاز کامل، در غیر این صورت ۰
    return 5.0 if n == 0 else 0.0

def score_unreal_week(good_ratio: float, weight: float = 10.0) -> float:
    """
    وعدهٔ غیرواقعی (Training - 10%):
    اگر حداقل ۹۵٪ روزها مقدار ستون 'وعده غیرواقعی' برابر ۱ باشد → امتیاز کامل
    در غیر این صورت → ۰
    """
    if good_ratio is None:
        return 0.0
    return weight if good_ratio >= 0.95 else 0.0


def compute_total_qc_score_from_row(r: pd.Series) -> float:
    """
    r: یک ردیف از stats (خروجی compute_weekly_stats)
    ستون‌های استفاده‌شده:

      متریک‌های قبلی:
        - "خوداظهاری اشتباه"
        - "ساعات طلایی"
        - "ساعات کاری کل"
        - "روزهای کاری"
        - "مشارکت با تاخیر"
        - "مشارکت‌های ناقص"
        - "تعداد فیلد های تخصیص داده شده"

      متریک‌های جدید:
        - "banner_daily"
        - "bannerQuality_daily"
        - "presenceQuality_daily"
        - "نمره حضور توسط راننده ها"
        - "نمره آموزش توسط راننده ها"
        - "وعده ی غیرواقعی"
        - "تطبیق"
        - "فعال سازی"
    """

    # -------- متریک‌های قبلی --------
    assigned   = r.get("تعداد فیلد های تخصیص داده شده", 0) or 0
    work_days  = r.get("روزهای کاری", 0) or 0

    false_count = r.get("خوداظهاری اشتباه", 0) or 0
    shift_delay = r.get("مشارکت با تاخیر", 0) or 0
    incomplete  = r.get("مشارکت‌های ناقص", 0) or 0

    # golden_val = تعداد روزهایی که goldentime = 1 بوده
    golden_val  = r.get("ساعات طلایی", 0) or 0

    total_hours = r.get("ساعات کاری کل", 0) or 0

    if work_days:
        zero_golden_count = max(float(work_days) - float(golden_val), 0.0)
    else:
        zero_golden_count = 0.0


    # درصد خوداظهاری اشتباه
    if assigned and assigned != 0:
        pct_false = 100.0 * float(false_count) / float(assigned)
    else:
        pct_false = 0.0


    # میانگین ساعت کاری در روز
    if work_days and work_days != 0:
        avg_hours = float(total_hours) / float(work_days)
    else:
        avg_hours = 0.0

    s_false = score_false_check(pct_false)           # max 20
    s_gold  = score_goldentime(zero_golden_count)    # max 7
    s_hrs   = score_total_work_hours(avg_hours)      # max 8
    s_sdl   = score_shift_delay(shift_delay)         # max 8
    s_inc   = score_incomplete(incomplete)           # max 2

    # Absence (روزهای بدون مشارکت) → max 5
    # این متریک جدید اضافه شده است تا بخش غیبت در TOTAL QC SCORE لحاظ شود.
    absence_val = r.get("روزهای بدون مشارکت", 0) or 0
    s_abs   = score_absence_total(absence_val)       # max 5

    # -------- متریک‌های جدید (QC راننده + QC ناظر + Tatbigh + Activation) --------

    # ۱) ستون‌های ناظر میدانی
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    # banner_daily همان bannerQuality_daily است
    banner_qual_val = r.get("bannerQuality_daily")
    if banner_qual_val is not None and not (hasattr(pd, "isna") and pd.isna(banner_qual_val)):
        banner_qual_pct = _to_pct_0_100(banner_qual_val)
    else:
        banner_qual_pct = None
    banner_daily_pct = banner_qual_pct  # banner_daily همان bannerQuality_daily است
    
    presence_qual_val = r.get("presenceQuality_daily")
    if presence_qual_val is not None and not (hasattr(pd, "isna") and pd.isna(presence_qual_val)):
        presence_qual_pct = _to_pct_0_100(presence_qual_val)
    else:
        presence_qual_pct = None

    s_banner_daily  = score_banner_daily(banner_daily_pct)        # max 5
    s_banner_qual   = score_banner_quality(banner_qual_pct)       # max 5
    s_presence_qual = score_presence_quality(presence_qual_pct)   # max 5

    # ۲) ستون‌های QC راننده‌ها
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    pres_score_val = r.get("نمره حضور توسط راننده ها")
    if pres_score_val is not None and not (hasattr(pd, "isna") and pd.isna(pres_score_val)):
        pres_score_pct = _to_pct_0_100(pres_score_val)
    else:
        pres_score_pct = None
    
    edu_score_val = r.get("نمره آموزش توسط راننده ها")
    if edu_score_val is not None and not (hasattr(pd, "isna") and pd.isna(edu_score_val)):
        edu_score_pct = _to_pct_0_100(edu_score_val)
    else:
        edu_score_pct = None

    # اگر هر دو نمره QC راننده خالی بودند → دیفالت "متوسط" = 70%
    if (pres_score_pct is None) and (edu_score_pct is None):
        pres_score_pct = 70.0
        edu_score_pct  = 70.0

    s_pres_score = score_presence_score(pres_score_pct)   # max 5
    s_edu_score  = score_education_score(edu_score_pct)   # max 10

    # --- وعده غیرواقعی (unreal) ---
    unreal_val = r.get("وعده ی غیرواقعی")
    
    # اگر مقدار خالی بود → امتیاز کامل (همانند compute_total_qc_score_details)
    if unreal_val is None or (hasattr(pd, "isna") and pd.isna(unreal_val)):
        s_unreal = 10.0
    else:
        unreal_good = float(unreal_val)  # تعداد دفعاتی که =1 بوده
        # نسبت روزهای "خوب" = 1
        if work_days and work_days > 0:
            unreal_ratio = unreal_good / float(work_days)
        else:
            unreal_ratio = None
        s_unreal = score_unreal_week(unreal_ratio, weight=10.0)


    # ۴) Tatbigh (تطبیق) – درصد نقص مدارک
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    tatbigh_val = r.get("تطبیق")
    if tatbigh_val is not None and not (hasattr(pd, "isna") and pd.isna(tatbigh_val)):
        tatbigh_pct = _to_pct_0_100(tatbigh_val)
    else:
        tatbigh_pct = None
    s_tatbigh   = score_tatbigh(tatbigh_pct)                      # max 5

    # ۵) Activation – درصد راننده‌های فعال
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    # ابتدا "درصد راننده‌های فعال" را چک کن، اگر نبود "فعال سازی" را چک کن
    activation_val = r.get("درصد راننده‌های فعال")
    if activation_val is None or (hasattr(pd, "isna") and pd.isna(activation_val)):
        activation_val = r.get("فعال سازی")
    if activation_val is not None and not (hasattr(pd, "isna") and pd.isna(activation_val)):
        activation_pct = _to_pct_0_100(activation_val)
    else:
        activation_pct = None
    s_activation   = score_activation(activation_pct)             # max 5

    # -------- جمع کل --------
    # جمع کل امتیازها (شامل غیبت جدید):
    raw = (
        s_false + s_gold + s_hrs + s_sdl + s_inc + s_abs +    # 50 قدیمی + غیبت
        s_banner_daily + s_banner_qual + s_presence_qual +    # 5+5+5
        s_pres_score + s_edu_score +                          # 5+10
        s_unreal + s_tatbigh + s_activation                    # 10+5+5
    )

    # حداکثر امتیاز تئوریک:
    # متریک‌های قدیمی: 20 + 7 + 8 + 8 + 2 + 5 = 50
    # متریک‌های جدید: 5 + 5 + 5 + 5 + 10 + 10 + 5 + 5 = 50
    # جمع کل = 100
    max_raw = 100.0

    if max_raw <= 0:
        return 0.0
    return round((raw / max_raw) * 100.0, 1)


def compute_total_qc_score_details(r: pd.Series) -> dict:
    """
    محاسبه جزئیات TOTAL QC SCORE برای نمایش در داشبورد
    خروجی: دیکشنری با جزئیات هر متریک و امتیاز آن
    """
    details = {}
    
    # -------- متریک‌های قبلی --------
    assigned   = r.get("تعداد فیلد های تخصیص داده شده", 0) or 0
    work_days  = r.get("روزهای کاری", 0) or 0
    false_count = r.get("خوداظهاری اشتباه", 0) or 0
    shift_delay = r.get("مشارکت با تاخیر", 0) or 0
    incomplete  = r.get("مشارکت‌های ناقص", 0) or 0
    golden_val  = r.get("ساعات طلایی", 0) or 0
    total_hours = r.get("ساعات کاری کل", 0) or 0

    # درصد خوداظهاری اشتباه
    if assigned and assigned != 0:
        pct_false = 100.0 * float(false_count) / float(assigned)
    else:
        pct_false = 0.0

    # تعداد دفعاتی که گلدن‌تایم صفر بوده = روزهای کاری - روزهای با گلدن‌تایم
    if work_days:
        zero_golden_count = max(float(work_days) - float(golden_val), 0.0) if work_days else 0.0
    else:
        zero_golden_count = 0.0

    # میانگین ساعت کاری در روز
    if work_days and work_days != 0:
        avg_hours = float(total_hours) / float(work_days)
    else:
        avg_hours = 0.0

    s_false = score_false_check(pct_false)
    s_gold  = score_goldentime(zero_golden_count)
    s_hrs   = score_total_work_hours(avg_hours)
    s_sdl   = score_shift_delay(shift_delay)
    s_inc   = score_incomplete(incomplete)

    details["خوداظهاری اشتباه"] = {
        "مقدار": f"{pct_false:.1f}%",
        "امتیاز": f"{s_false:.1f}",
        "حداکثر": "20"
    }
    details["ساعات طلایی"] = {
        "مقدار": f"{golden_val:.0f}",          # 👈 مقدار واقعی: چند روز گلدن بوده
        "امتیاز": f"{s_gold:.1f}",             # 👈 امتیاز بر اساس تعداد صفرها
        "حداکثر": "7"
    }

    details["میانگین ساعات کاری"] = {
        "مقدار": f"{avg_hours:.2f} ساعت/روز",
        "امتیاز": f"{s_hrs:.1f}",
        "حداکثر": "8"
    }
    details["مشارکت با تاخیر"] = {
        "مقدار": f"{shift_delay:.0f}",
        "امتیاز": f"{s_sdl:.1f}",
        "حداکثر": "8"
    }
    details["مشارکت‌های ناقص"] = {
        "مقدار": f"{incomplete:.0f}",
        "امتیاز": f"{s_inc:.1f}",
        "حداکثر": "2"
    }

    # --- غیبت (روزهای بدون مشارکت) ---
    # این متریک جدید برای TOTAL QC SCORE اضافه شده است.
    # اگر absence_val برابر ۰ باشد امتیاز کامل (۵)، در غیر این صورت امتیاز صفر می‌گیرد.
    absence_val = r.get("روزهای بدون مشارکت", 0) or 0
    try:
        _absence_display = float(absence_val)
    except Exception:
        _absence_display = 0.0
    # امتیاز غیبت
    s_abs = score_absence_total(absence_val)
    details["روزهای بدون مشارکت"] = {
        "مقدار": f"{_absence_display:.0f}",
        "امتیاز": f"{s_abs:.1f}",
        "حداکثر": "5"
    }

    # -------- متریک‌های جدید --------
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    # banner_daily همان bannerQuality_daily است
    banner_qual_val = r.get("bannerQuality_daily")
    if banner_qual_val is not None and not (hasattr(pd, "isna") and pd.isna(banner_qual_val)):
        banner_qual_pct = _to_pct_0_100(banner_qual_val)
    else:
        banner_qual_pct = None
    banner_daily_pct = banner_qual_pct  # banner_daily همان bannerQuality_daily است
    
    presence_qual_val = r.get("presenceQuality_daily")
    if presence_qual_val is not None and not (hasattr(pd, "isna") and pd.isna(presence_qual_val)):
        presence_qual_pct = _to_pct_0_100(presence_qual_val)
    else:
        presence_qual_pct = None

    s_banner_daily  = score_banner_daily(banner_daily_pct)
    s_banner_qual   = score_banner_quality(banner_qual_pct)
    s_presence_qual = score_presence_quality(presence_qual_pct)

    details["banner_daily"] = {
        "مقدار": f"{banner_daily_pct:.1f}%" if banner_daily_pct is not None else "ندارد",
        "امتیاز": f"{s_banner_daily:.1f}",
        "حداکثر": "5"
    }
    details["کیفیت بنر (bannerQuality)"] = {
        "مقدار": f"{banner_qual_pct:.1f}%" if banner_qual_pct is not None else "ندارد",
        "امتیاز": f"{s_banner_qual:.1f}",
        "حداکثر": "5"
    }
    details["کیفیت حضور میدانی (presenceQuality)"] = {
        "مقدار": f"{presence_qual_pct:.1f}%" if presence_qual_pct is not None else "ندارد",
        "امتیاز": f"{s_presence_qual:.1f}",
        "حداکثر": "5"
    }

    # QC راننده‌ها
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    pres_score_val = r.get("نمره حضور توسط راننده ها")
    if pres_score_val is not None and not (hasattr(pd, "isna") and pd.isna(pres_score_val)):
        pres_score_pct = _to_pct_0_100(pres_score_val)
    else:
        pres_score_pct = None
    
    edu_score_val = r.get("نمره آموزش توسط راننده ها")
    if edu_score_val is not None and not (hasattr(pd, "isna") and pd.isna(edu_score_val)):
        edu_score_pct = _to_pct_0_100(edu_score_val)
    else:
        edu_score_pct = None

    if (pres_score_pct is None) and (edu_score_pct is None):
        pres_score_pct = 70.0
        edu_score_pct  = 70.0

    s_pres_score = score_presence_score(pres_score_pct)
    s_edu_score  = score_education_score(edu_score_pct)

    details["نمره حضور توسط راننده‌ها"] = {
        "مقدار": f"{pres_score_pct:.1f}%" if pres_score_pct is not None else "ندارد",
        "امتیاز": f"{s_pres_score:.1f}",
        "حداکثر": "5"
    }
    details["نمره آموزش توسط راننده‌ها"] = {
        "مقدار": f"{edu_score_pct:.1f}%" if edu_score_pct is not None else "ندارد",
        "امتیاز": f"{s_edu_score:.1f}",
        "حداکثر": "10"
    }

    # --- وعده غیرواقعی (unreal) ---
    unreal_val = r.get("وعده ی غیرواقعی")

    # اگر مقدار خالی بود → امتیاز کامل
    if unreal_val is None or (hasattr(pd, "isna") and pd.isna(unreal_val)):
        s_unreal = 10.0
        unreal_ratio = None
    else:
        unreal_good = float(unreal_val)
        if work_days and work_days > 0:
            unreal_ratio = unreal_good / float(work_days)
        else:
            unreal_ratio = None
        s_unreal = score_unreal_week(unreal_ratio, weight=10.0)

    details["وعده غیرواقعی"] = {
        "مقدار": f"{(unreal_ratio*100):.1f}%" if unreal_ratio is not None else "-",
        "امتیاز": f"{s_unreal:.1f}",
        "حداکثر": "10"
    }


    # Tatbigh
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    tatbigh_val = r.get("تطبیق")
    if tatbigh_val is not None and not (hasattr(pd, "isna") and pd.isna(tatbigh_val)):
        tatbigh_pct = _to_pct_0_100(tatbigh_val)
    else:
        tatbigh_pct = None
    s_tatbigh   = score_tatbigh(tatbigh_pct)

    details["تطبیق (نقص مدارک)"] = {
        "مقدار": f"{tatbigh_pct:.1f}%" if tatbigh_pct is not None else "ندارد",
        "امتیاز": f"{s_tatbigh:.1f}",
        "حداکثر": "5"
    }

    # Activation
    # اگر مقدار خالی بود، None بذار (نه 0) تا در میانگین نیفته
    # ابتدا "درصد راننده‌های فعال" را چک کن، اگر نبود "فعال سازی" را چک کن
    activation_val = r.get("درصد راننده‌های فعال")
    if activation_val is None or (hasattr(pd, "isna") and pd.isna(activation_val)):
        activation_val = r.get("فعال سازی")
    if activation_val is not None and not (hasattr(pd, "isna") and pd.isna(activation_val)):
        activation_pct = _to_pct_0_100(activation_val)
    else:
        activation_pct = None
    s_activation   = score_activation(activation_pct)

    details["فعال‌سازی"] = {
        "مقدار": f"{activation_pct:.1f}%" if activation_pct is not None else "ندارد",
        "امتیاز": f"{s_activation:.1f}",
        "حداکثر": "5"
    }

    # -------- جمع کل --------
    # مجموع امتیازات شامل متریک غیبت
    raw = (
        s_false + s_gold + s_hrs + s_sdl + s_inc + s_abs +
        s_banner_daily + s_banner_qual + s_presence_qual +
        s_pres_score + s_edu_score +
        s_unreal + s_tatbigh + s_activation
    )

    # حداکثر امتیاز تئوریک با احتساب غیبت: ۵۰ (قدیمی + غیبت) + ۵۰ (جدید) = ۱۰۰
    max_raw = 100.0
    final_score = round((raw / max_raw) * 100.0, 1) if max_raw > 0 else 0.0

    # خلاصهٔ امتیازها
    details["_summary"] = {
        "جمع امتیازها": f"{raw:.1f}",
        "حداکثر امتیاز": "100",
        "نمره نهایی": f"{final_score:.1f}"
    }

    return details


def _safe_float(v, default=0.0):
    try:
        if v is None or (hasattr(pd, "isna") and pd.isna(v)):
            return default
        return float(v)
    except Exception:
        return default

def _absence_good(n):
    """۰ → عالی، ۱–۲ → متوسط، بیشتر → ضعیف (۰–۱)"""
    x = _safe_float(n, 0.0)
    if x == 0:
        return 1.0
    elif x <= 2:
        return 0.5
    else:
        return 0.0

def _qc_block_weighted_from_row(r: pd.Series) -> float:
    """
    زیر‌امتیاز QC (۰ تا ۲۰) بر اساس وزن‌هایی که گفتی:
      حضور راننده: 2.5
      آموزش راننده: 2.5
      QC ناظر میدانی (presenceQuality/banner...) مجموعاً: 5
      false_check: 2
      goldentime: 2
      totalWorkHours: 2
      shift_delay: 1.5
      incomplete: 1
      absence: 1.5
    """
    # --- نمره‌های درصدی راننده / ناظر (۰–۱) ---
    drv_pres = _percent_to_0_1(r.get("نمره حضور توسط راننده ها"))
    drv_edu  = _percent_to_0_1(r.get("نمره آموزش توسط راننده ها"))
    field_q  = _percent_to_0_1(r.get("نمره QC ناظران میدانی"))

    # 2.5 + 2.5 + 5 (برای presenceQuality_daily + banner_daily + bannerQuality_daily به‌صورت تجمیعی)
    qc_part = (
        2.5 * (drv_pres if not math.isnan(drv_pres) else 0.0) +
        2.5 * (drv_edu  if not math.isnan(drv_edu)  else 0.0) +
        5.0 * (field_q  if not math.isnan(field_q)  else 0.0)
    )

    # --- متریک‌های رفتاری با منطق آستانه‌ای قبلی (نرمال‌شده به ۰–۱) ---
    assigned   = _safe_float(r.get("تعداد فیلد های تخصیص داده شده"), 0.0)
    work_days  = _safe_float(r.get("روزهای کاری"), 0.0)
    false_cnt  = _safe_float(r.get("خوداظهاری اشتباه"), 0.0)
    shift_d    = _safe_float(r.get("مشارکت با تاخیر"), 0.0)
    incomplete = _safe_float(r.get("مشارکت‌های ناقص"), 0.0)
    golden_val = _safe_float(r.get("ساعات طلایی"), 0.0)
    total_hrs  = _safe_float(r.get("ساعات کاری کل"), 0.0)
    absence    = _safe_float(r.get("روزهای بدون مشارکت"), 0.0)

    # درصد خوداظهاری اشتباه
    if assigned > 0:
        pct_false = 100.0 * false_cnt / assigned
    else:
        pct_false = 0.0

    # میانگین ساعت
    if work_days > 0:
        avg_hours = total_hrs / work_days
    else:
        avg_hours = 0.0

    zero_golden_count = max(float(work_days) - float(golden_val), 0.0) if work_days else 0.0

    fc_good   = score_false_check(pct_false)      / 20.0   # ۰–۱
    gold_good = score_goldentime(zero_golden_count) / 7.0
    hrs_good  = score_total_work_hours(avg_hours) / 8.0
    sdl_good  = score_shift_delay(shift_d)        / 8.0
    inc_good  = score_incomplete(incomplete)      / 2.0
    abs_good  = _absence_good(absence)           # ۰–۱

    qc_part += 2.0  * fc_good
    qc_part += 2.0  * gold_good
    qc_part += 2.0  * hrs_good
    qc_part += 1.5  * sdl_good
    qc_part += 1.0  * inc_good
    qc_part += 1.5  * abs_good

    # سقف تئوریک ۲۰ است
    return qc_part


def _attach_weighted_score(df: pd.DataFrame) -> pd.DataFrame:
    """
    روی df ستون Score را اضافه می‌کند بر اساس:
      80% → میانگین جذب روزانه
      20% → بلوک QC (۰–۲۰ نرمال‌شده)
    """
    df = df.copy()

    # --- ۱) میانگین جذب روزانه (تعداد جذب / روزهای کاری) ---
    acq = pd.to_numeric(df.get("تعداد جذب"), errors="coerce")
    wd  = pd.to_numeric(df.get("روزهای کاری"), errors="coerce")
    wd  = wd.replace(0, pd.NA)
    avg_acq = (acq / wd).fillna(0.0)

    # نرمال‌سازی ۰–۱ بین همین گروه
    acq_norm = _minmax(avg_acq).fillna(0.0)

    # --- ۲) QC بلوکی (۰–۲۰) → نرمال ۰–۱ ---
    qc_raw  = df.apply(_qc_block_weighted_from_row, axis=1)
    qc_norm = (qc_raw / 20.0).fillna(0.0)

    df["Score"] = 0.8 * acq_norm + 0.2 * qc_norm
    return df


def build_top_performers(stats: pd.DataFrame, city: str | None, top_n: int = 10) -> pd.DataFrame:
    # اول امتیاز رتبه را بر اساس منطق جدید محاسبه می‌کنیم
    df = compute_agent_rank_score(stats)
    if city and city != "همه":
        df = df[df["شهر"] == city].copy()
    if df.empty:
        return df

    cols_show = [
        "آیدی","نام و نام خانوادگی","شهر","سرپرست","تیم‌ لید",
        "تعداد جذب","روزهای کاری","روزهای بدون مشارکت",
        "نمره QC راننده ها","نمره QC ناظران میدانی","TOTAL QC SCORE","Score"
    ]
    cols_show = [c for c in cols_show if c in df.columns]
    out = df[cols_show].copy().sort_values("Score", ascending=False).head(top_n).reset_index(drop=True)

    if "نمره QC ناظران میدانی" in out.columns:
        out["نمره QC ناظران میدانی"] = out["نمره QC ناظران میدانی"].apply(lambda v: fmt_val(v, percent=True))
    out["Score"] = out["Score"].apply(lambda v: f"{v:.3f}")
    return out


def compute_city_scores_ranks(stats: pd.DataFrame) -> pd.DataFrame:
    df = compute_agent_rank_score(stats)
    if df.empty or "شهر" not in df.columns:
        df["CityRank"] = None
        df["CityCount"] = None
        return df

    pieces = []
    for city, g in df.groupby("شهر", dropna=False):
        gg = g.copy().sort_values("Score", ascending=False).reset_index(drop=True)
        gg["CityRank"] = range(1, len(gg) + 1)
        gg["CityCount"] = gg["آیدی"].nunique()
        pieces.append(gg)

    out = pd.concat(pieces, axis=0).reset_index(drop=True)
    return out



def city_rank_for_id(stats: pd.DataFrame, agent_id: str | int) -> str | None:
    df = compute_city_scores_ranks(stats)
    row = df[df["آیدی"].astype(str) == str(agent_id)]
    if row.empty:
        return None
    r = row.iloc[0]
    try:
        x = int(r["CityRank"]); y = int(r["CityCount"])
        return f"{x} از {y}"
    except Exception:
        return None


def compute_team_scores_ranks(stats: pd.DataFrame) -> pd.DataFrame:
    df = compute_agent_rank_score(stats)
    # ستون سرپرست ممکن است با کاراکترهای عجیبی ذخیره شده باشد؛ سعی می‌کنیم هر دو را پوشش دهیم
    col_sup = "سرپرست"
    if col_sup not in df.columns and "सरपरست" in df.columns:
        col_sup = "सरपरست"

    if df.empty or col_sup not in df.columns:
        df["TeamRank"] = None
        df["TeamCount"] = None
        return df

    pieces = []
    for sup, g in df.groupby(col_sup, dropna=False):
        gg = g.copy().sort_values("Score", ascending=False).reset_index(drop=True)
        gg["TeamRank"] = range(1, len(gg) + 1)
        gg["TeamCount"] = gg["آیدی"].nunique()
        pieces.append(gg)

    out = pd.concat(pieces, axis=0).reset_index(drop=True)
    return out



def team_rank_for_id(stats: pd.DataFrame, agent_id: str | int) -> str | None:
    df = compute_team_scores_ranks(stats)
    row = df[df["آیدی"].astype(str) == str(agent_id)]
    if row.empty:
        return None
    r = row.iloc[0]
    try:
        x = int(r["TeamRank"]); y = int(r["TeamCount"])
        return f"{x} از {y}"
    except Exception:
        return None


# ---------- UI ----------
st.set_page_config(page_title="کارنامه هفتگی بازاریابان و سرپرستان", page_icon="📊", layout="centered")

st.markdown(
    """
<div class="sup-wrap" style="max-width:1280px;margin:0 auto 8px;">
  <div style="background: linear-gradient(135deg,#1e3a8a 0%, #0ea5e9 60%, #22d3ee 100%); color:#fff; border-radius:16px;padding: 18px; box-shadow:var(--shadow);">
    <div style="font-size:24px; font-weight:800; margin:0 0 4px 0;">کارنامه هفتگی بازاریابان و سرپرستان</div>
    <div style="opacity:.95; font-size:13px">خلاصهٔ عملکرد ۷ روز اخیر با امکان دانلود خروجی HTML برای افراد و سرپرستان</div>
  </div>
</div>
""",
    unsafe_allow_html=True
)

with st.sidebar:
    st.markdown("### تنظیمات برند")
    logo_file = st.file_uploader("لوگو (PNG/JPG/SVG)", type=["png","jpg","jpeg","svg"])
logo_b64 = image_to_b64(logo_file)

uploaded = st.file_uploader("فایل اکسل را انتخاب کنید", type=["xlsx"])

if uploaded:
    df_raw = load_excel(uploaded)
    mp = find_column_mapping(df_raw)

    # --- خروجی QC برای مهر ---
    qc_mehr = export_qc_last_record_for_month(df_raw, mp, 7, 1404, "QC_Mehr_1404.xlsx")

    # --- خروجی QC برای آبان ---
    qc_aban = export_qc_last_record_for_month(df_raw, mp, 8, 1404, "QC_Aban_1404.xlsx")

    # نمایش در Streamlit
    st.write("📌 QC مهر 1404:")
    st.dataframe(qc_mehr)

    st.write("📌 QC آبان 1404:")
    st.dataframe(qc_aban)

    # باقی کد شما
    dates = pd.to_datetime(df_raw[mp["date"]], errors="coerce").dropna()
    default_anchor = dates.max().date() if not dates.empty else datetime.today().date()

    col1, col2 = st.columns([1,1], gap="medium")
    with col1:
        st.write("Anchor date (7-day window):")
        anchor = st.date_input("Anchor date", value=default_anchor)
    with col2:
        mode = st.radio("نمایش:", ["کارت بازاریاب", "گزارش سرپرست"], horizontal=True)

    # ✅ مبنای تاریخ
    basis = st.radio("مبنای تاریخ:", ["آخرین هفتهٔ موجود", "Anchor انتخابی"], horizontal=True)

    # === ساخت df7 بر اساس مبنا ===
    if basis == "Anchor انتخابی":
        # یک پنجرهٔ ۷روزه حول Anchor برای همه
        df_raw["_date"] = pd.to_datetime(df_raw[mp["date"]], errors="coerce").dt.date
        df7 = filter_last_7_days(
            df_raw,
            mp["date"],
            anchor_date=datetime.combine(anchor, datetime.min.time())
        )
    else:
        # آخرین هفتهٔ موجود برای هر شهر (حداکثر تاریخ هر شهر → ۶ روز قبلش)
        df_raw["_date"] = pd.to_datetime(df_raw[mp["date"]], errors="coerce").dt.date
        if mp.get("city"):
            df7 = filter_last_7_days_per_group(df_raw, mp["date"], mp["city"])
        else:
            df7 = filter_last_7_days(df_raw, mp["date"])

    if df7.empty:
        st.warning("برای این بازهٔ ۷ روزه داده‌ای پیدا نشد. مبنای تاریخ یا Anchor را تغییر دهید.")
        st.stop()

    if mp.get("city"):
        unique_cities = df7[mp["city"]].dropna().unique()
        st.info(f"تعداد شهرهای منحصر به فرد در ۷ روز اخیر: **{len(unique_cities)}**")
        st.caption(f"شهرهای یافت شده: {', '.join(map(str, unique_cities))}")

    # === تاریخ مرجع یکسان برای همهٔ خروجی‌ها از خودِ df7 ===
    _from_g_ref = pd.to_datetime(pd.Series(df7["_date"])).min().date()
    _to_g_ref   = pd.to_datetime(pd.Series(df7["_date"])).max().date()
    from_j_ref  = to_jalali_words(_from_g_ref, persian_digits=True)
    to_j_ref    = to_jalali_words(_to_g_ref,   persian_digits=True)

    # === مرجع یکسان: آخرین هفته‌ی موجود در کل داده (week max) ===
    if mp.get("week") and mp["week"] in df7.columns:
        _wk_series = pd.to_numeric(df7[mp["week"]], errors="coerce")
        max_week = _wk_series.max()
        df_week_latest = df7[_wk_series == max_week].copy()

        # تاریخ مرجع = آخرین روزِ دیتای همین هفته
        _to_g_ref = pd.to_datetime(df_week_latest["_date"]).max().date()
        # همیشه بازه‌ی ۷ روزه نشان بده، حتی اگر فقط یک روز داده داریم
        _from_g_ref = _to_g_ref - timedelta(days=6)
    else:
        # fallback اگر ستون week نداشتیم
        _to_g_ref   = pd.to_datetime(pd.Series(df7["_date"])).max().date()
        _from_g_ref = _to_g_ref - timedelta(days=6)

    from_j_ref = to_jalali_words(_from_g_ref, persian_digits=True)
    to_j_ref   = to_jalali_words(_to_g_ref,   persian_digits=True)

    # ======== MODE 1: Agent Cards ========
    if mode == "کارت بازاریاب":

        st.write("MIN–MAX df7:", 
                 pd.to_datetime(pd.Series(df7["_date"])).min(), 
                 pd.to_datetime(pd.Series(df7["_date"])).max())

        # بازه کلی فقط برای Top10 و Batch
        _from_g = pd.to_datetime(pd.Series(df7["_date"])).min().date()
        _to_g   = pd.to_datetime(pd.Series(df7["_date"])).max().date()
        from_j = to_jalali_words(_from_g, persian_digits=True)
        to_j   = to_jalali_words(_to_g,   persian_digits=True)

        stats = compute_weekly_stats(df7, mp)

        # ✅ میانگین کلی TOTAL QC SCORE در همین بازه
        mean_qc = None
        if "TOTAL QC SCORE" in stats.columns:
            vals_qc = pd.to_numeric(stats["TOTAL QC SCORE"], errors="coerce")
            mean_qc = vals_qc.mean()

        if mean_qc is not None and not pd.isna(mean_qc):
            st.info(f"میانگین TOTAL QC SCORE این بازه: **{mean_qc:.1f}**")

        with st.expander("خلاصهٔ بازاریاب‌ها (اختیاری)", expanded=False):
            st.dataframe(stats, use_container_width=True)


        # --- Top 10 – تنظیمات (اختیاری) ---
        with st.expander("Top 10 – تنظیمات (اختیاری)", expanded=False):
            tp_col1, tp_col2 = st.columns([1,1])
            with tp_col1:
                scope = st.radio("دامنه:", ["کشور", "شهر"], horizontal=True)
                city_choice = "همه" if scope == "کشور" else st.selectbox("شهر", sorted(stats["شهر"].dropna().unique()))
            with tp_col2:
                top_n = st.number_input("تعداد (N)", min_value=1, max_value=100, value=10, step=1)

        top_df = build_top_performers(stats, city=city_choice, top_n=int(top_n))
        if not top_df.empty:
            # ✅ تاریخ هدر Top10 = مرجع واحد (week آخر)
            top_html = build_top10_html(city_choice, top_df, from_j_ref, to_j_ref, logo_src=logo_b64)
            html_download_button(
                f"top10_{(city_choice if city_choice!='همه' else 'all_cities')}.html".replace(" ", "_"),
                top_html,
                "دانلود HTML Top 10"
            )

        spacer(12)
        st.markdown("---")
        st.subheader("کارت تکی")
        ids = list(stats["آیدی"].astype(str).unique())
        sel_id = st.selectbox("آیدی:", ids)
        person = stats[stats["آیدی"].astype(str) == str(sel_id)].iloc[0].to_dict()  
        
        # دیباگ: بررسی mapping و داده‌ها
        with st.expander("🔍 دیباگ روند جذب (برای بررسی مشکل)", expanded=False):
            st.write("**Mapping ستون‌ها:**")
            st.write(f"- week column: `{mp.get('week', 'NOT FOUND')}`")
            st.write(f"- acq column: `{mp.get('acq', 'NOT FOUND')}`")
            st.write(f"- id column: `{mp.get('id', 'NOT FOUND')}`")
            
            if mp.get("week") and mp.get("acq"):
                # بررسی نوع داده ستون acq
                acq_col = mp["acq"]
                st.write(f"\n**نوع داده ستون '{acq_col}':**")
                st.write(f"- dtype: {df_raw[acq_col].dtype}")
                st.write(f"- نمونه مقادیر (اولین 5 مقدار غیر None):")
                sample_vals = df_raw[acq_col].dropna().head(5).tolist()
                st.write(sample_vals if sample_vals else "همه None هستند!")
                
                # بررسی داده‌های این آیدی در df_raw
                debug_df = df_raw[[mp["id"], mp["week"], mp["acq"]]].copy()
                debug_df = debug_df[debug_df[mp["id"]].astype(str).str.strip() == str(sel_id).strip()]
                st.write(f"\n**تعداد ردیف‌های این آیدی در df_raw: {len(debug_df)}**")
                
                # بررسی داده‌های این آیدی در df7 (7 روز اخیر)
                if mp["acq"] in df7.columns:
                    debug_df7 = df7[df7[mp["id"]].astype(str).str.strip() == str(sel_id).strip()]
                    st.write(f"**تعداد ردیف‌های این آیدی در df7 (7 روز اخیر): {len(debug_df7)}**")
                    if not debug_df7.empty:
                        st.write(f"**نمونه داده‌های df7 (اولین 5 ردیف):**")
                        st.dataframe(debug_df7[[mp["id"], mp["week"], mp["acq"]]].head(5), use_container_width=True)
                        # بررسی اینکه آیا daily_acq خالی است
                        non_null_acq = debug_df7[mp["acq"]].notna().sum()
                        st.write(f"**تعداد ردیف‌های با daily_acq غیرخالی در df7: {non_null_acq} از {len(debug_df7)}**")
                
                if not debug_df.empty:
                    st.write("\n**نمونه داده‌ها از df_raw (اولین 10 ردیف):**")
                    st.dataframe(debug_df.head(10), use_container_width=True)
                    
                    # بررسی اینکه آیا daily_acq خالی است
                    non_null_count = debug_df[mp["acq"]].notna().sum()
                    st.write(f"**تعداد ردیف‌های با daily_acq غیرخالی در df_raw: {non_null_count} از {len(debug_df)}**")
                    
                    # بررسی هفته‌ها
                    debug_df["_week"] = to_numeric_clean(debug_df[mp["week"]])
                    debug_df["_acq"] = to_numeric_clean(debug_df[mp["acq"]])
                    
                    # فقط ردیف‌هایی که هم week و هم acq معتبر دارند
                    debug_df_valid = debug_df.dropna(subset=["_week", "_acq"])
                    st.write(f"**تعداد ردیف‌های با week و acq معتبر: {len(debug_df_valid)}**")
                    
                    if not debug_df_valid.empty:
                        weeks_unique = sorted(debug_df_valid["_week"].unique())
                        st.write(f"\n**هفته‌های موجود (با acq معتبر): {weeks_unique}**")
                        st.write(f"**آخرین 4 هفته: {weeks_unique[-4:] if len(weeks_unique) >= 4 else weeks_unique}**")
                        
                        # جمع هر هفته
                        agg_debug = debug_df_valid.groupby("_week")["_acq"].sum().reset_index()
                        agg_debug = agg_debug.sort_values("_week")
                        st.write("\n**جمع جذب هر هفته (فقط ردیف‌های معتبر):**")
                        st.dataframe(agg_debug, use_container_width=True)
                    else:
                        st.warning("⚠️ بعد از تبدیل به عدد، هیچ ردیفی با week و acq معتبر باقی نماند!")
                else:
                    st.warning(f"⚠️ برای آیدی {sel_id} هیچ داده‌ای در df_raw پیدا نشد!")
            else:
                st.error("❌ ستون week یا acq پیدا نشد!")
        
        # تلاش اول: استفاده از df_raw (همه داده‌ها)
        #
        # تعریف یک متغیر برای مشخص کردن دیتافریم منبع برای ترند جذب و در ادامه برای محاسبه
        # تعداد راننده‌های فعال. به‌صورت پیش‌فرض، df_raw در نظر گرفته می‌شود.
        df_for_trend = df_raw
        agg = build_acq_trend_last4weeks(df_for_trend, mp, sel_id)

        # اگر از df_raw نتیجه نگرفتیم، از df7 امتحان کنیم
        if (agg is None or agg.empty) and mp.get("week") and mp.get("acq"):
            with st.expander("⚠️ تلاش با df7 (7 روز اخیر)", expanded=False):
                st.info("از df_raw نتیجه نگرفتیم، در حال امتحان با df7...")
                # در صورت استفاده از df7، دیتافریم منبع ترند را به df7 تغییر می‌دهیم
                df_for_trend = df7
                agg = build_acq_trend_last4weeks(df_for_trend, mp, sel_id)
        
        if agg is not None and not agg.empty:
            st.info(f"✅ داده‌های ترند پیدا شد: {len(agg)} هفته")
            st.dataframe(agg, use_container_width=True)
        else:
            st.warning("⚠️ داده‌های ترند پیدا نشد یا خالی است!")
        
        chart_src = build_acq_trend_svg_from_df(agg, width=420, height=160, point_fs=14)

        # ✅ محاسبه تعداد راننده‌های فعال برای 4 هفته اخیر (همانند روند جذب)
        # در اینجا از df_for_trend استفاده می‌کنیم تا بازه‌ی محاسبه با همان دیتافریمی که
        # برای ترند جذب استفاده شده یکسان باشد (df_raw یا df7).
        _id_series_raw = df_for_trend[mp["id"]].astype(str).str.strip()
        df_agent_4weeks = df_for_trend[_id_series_raw == str(sel_id).strip()].copy()

        if not df_agent_4weeks.empty and mp.get("week"):
            # پیدا کردن 4 هفته اخیر بر اساس ستون week
            df_agent_4weeks["_week_num"] = to_numeric_clean(df_agent_4weeks[mp["week"]])
            df_agent_4weeks = df_agent_4weeks.dropna(subset=["_week_num"])
            
            if not df_agent_4weeks.empty:
                weeks_sorted = sorted(df_agent_4weeks["_week_num"].unique())
                if weeks_sorted:
                    last4_weeks = weeks_sorted[-4:]
                    df_agent_4weeks = df_agent_4weeks[df_agent_4weeks["_week_num"].isin(last4_weeks)]
                    
                    # محاسبه تعداد راننده‌های فعال و درصد برای 4 هفته اخیر
                    if mp.get("active_drivers") and mp["active_drivers"] in df_agent_4weeks.columns:
                        active_sum = to_numeric_clean(df_agent_4weeks[mp["active_drivers"]]).sum()
                        person["تعداد راننده‌های فعال"] = active_sum

                        if mp.get("acq") and mp["acq"] in df_agent_4weeks.columns:
                            acq_sum = to_numeric_clean(df_agent_4weeks[mp["acq"]]).sum()
                            if acq_sum > 0:
                                person["درصد راننده‌های فعال"] = (active_sum / acq_sum) * 100
                            else:
                                person["درصد راننده‌های فعال"] = None
                        else:
                            person["درصد راننده‌های فعال"] = None
                    else:
                        person["تعداد راننده‌های فعال"] = None
                        person["درصد راننده‌های فعال"] = None
                    
                    # ✅ محاسبه نقص مدارک: مجموع مقادیر هر ستون نقص مدارک طی ۴ هفته اخیر
                    doc_issue_cols_present = [col for col in DOC_ISSUE_COLS if col in df_agent_4weeks.columns]
                    if doc_issue_cols_present:
                        # جمع کل نقص‌ها در تمام ستون‌ها
                        total_doc_issues = 0.0
                        for col in doc_issue_cols_present:
                            try:
                                col_numeric = to_numeric_clean(df_agent_4weeks[col])
                                col_sum = col_numeric.sum()
                            except Exception:
                                col_sum = 0
                            # اگر NaN باشد، به‌عنوان ۰ در نظر بگیریم
                            if pd.isna(col_sum):
                                col_sum = 0
                            # ذخیره مقدار جمع‌شده برای هر ستون در person (برای بولت‌پوینت)
                            person[col] = col_sum
                            # فقط مقادیر مثبت را در جمع کل لحاظ می‌کنیم
                            try:
                                if float(col_sum) > 0:
                                    total_doc_issues += float(col_sum)
                            except Exception:
                                pass
                        # تعداد نقص مدارک = جمع مقادیر مثبت در تمام ستون‌ها
                        person["تعداد نقص مدارک"] = total_doc_issues if total_doc_issues > 0 else 0
                    else:
                        person["تعداد نقص مدارک"] = None

                    # ✅ محاسبه درصد نقص مدارک بر اساس نسبت تعداد نقص مدارک به مجموع جذب در همین ۴ هفته
                    # به جای استفاده از ستون Tatbigh، درصد نقص مدارک را خودمان محاسبه می‌کنیم.
                    try:
                        # محاسبه مجموع جذب این ۴ هفته
                        acq_sum_for_doc = None
                        if mp.get("acq") and mp["acq"] in df_agent_4weeks.columns:
                            acq_sum_for_doc = to_numeric_clean(df_agent_4weeks[mp["acq"]]).sum()
                        total_docs = person.get("تعداد نقص مدارک")
                        # اگر مقدار جذب و تعداد نقص مدارک موجود بود
                        if acq_sum_for_doc is not None and acq_sum_for_doc > 0 and total_docs is not None:
                            person["درصد نقص مدارک"] = (float(total_docs) / float(acq_sum_for_doc)) * 100
                        else:
                            # اگر داده جذب وجود نداشت یا صفر بود، درصورت وجود نقص مدارک درصد را صفر یا None قرار بده
                            if total_docs is not None and float(total_docs) == 0:
                                person["درصد نقص مدارک"] = 0
                            else:
                                person["درصد نقص مدارک"] = None
                    except Exception:
                        # در صورت خطا، درصد را None قرار بده
                        person["درصد نقص مدارک"] = None
                else:
                    person["تعداد راننده‌های فعال"] = None
                    person["درصد راننده‌های فعال"] = None
                    person["تعداد نقص مدارک"] = None
                    person["درصد نقص مدارک"] = None
            else:
                person["تعداد راننده‌های فعال"] = None
                person["درصد راننده‌های فعال"] = None
                person["تعداد نقص مدارک"] = None
                person["درصد نقص مدارک"] = None
        else:
            person["تعداد راننده‌های فعال"] = None
            person["درصد راننده‌های فعال"] = None
            person["تعداد نقص مدارک"] = None
            person["درصد نقص مدارک"] = None

        # فقط داده‌های همین آیدی در df7 برای نمایش تاریخ کارت (با strip برای مچ‌شدن امن)
        _id_series = df7[mp["id"]].astype(str).str.strip()
        df_agent = df7[_id_series == str(sel_id).strip()].copy()

        if df_agent.empty:
            st.warning("برای این آیدی در بازهٔ ۷ روز اخیر داده‌ای پیدا نشد.")
            st.stop()

        # ⛔️ مهم: بازهٔ ۷ روز اخیر را بر اساس max تاریخِ همین آیدی clamp کن (فقط برای نمایش تاریخ کارت)
        _dates_agent = pd.to_datetime(df_agent["_date"], errors="coerce")
        _to_g_sel = _dates_agent.max().date()
        _from_g_sel = (_to_g_sel - timedelta(days=6))

        from_j_sel = to_jalali_words(_from_g_sel, persian_digits=True)
        to_j_sel   = to_jalali_words(_to_g_sel,   persian_digits=True)

        rank_text = city_rank_for_id(stats, sel_id)
        team_rank_text = team_rank_for_id(stats, sel_id)

        chart_src = build_acq_trend_svg_uri(df_raw, mp, sel_id, width=420, height=160, point_fs=14)
        recent4 = acq_last_k_weeks_sum(df_raw, mp, sel_id, k=4)
        last_week_acq = acq_last_week_sum(df_raw, mp, sel_id)


        card_html = build_card_html(person, from_j_sel, to_j_sel, logo_src=logo_b64,
                            rank_text=rank_text, team_rank_text=team_rank_text, acq_chart_src=chart_src, acq_recent3_total=recent4, acq_last_week=last_week_acq)

        # هدر فقط بر اساس همین آیدی
        st.markdown(build_agent_date_block_shamsi(from_j_sel, to_j_sel, logo_src=logo_b64), unsafe_allow_html=True)
        spacer(8)

        # کارت + دانلود HTML با تاریخ درست همان آیدی
        card_html = build_card_html(person, from_j_sel, to_j_sel, logo_src=logo_b64, rank_text=rank_text, team_rank_text=team_rank_text, acq_chart_src=chart_src, acq_recent3_total=recent4, acq_last_week=last_week_acq)
        html_download_button(f"card_{person.get('آیدی','agent')}.html", card_html, "دانلود HTML همین کارت")
        st.components.v1.html(card_html, height=600, scrolling=True)

        # === جزئیات TOTAL QC SCORE ===
        spacer(12)
        with st.expander("📊 جزئیات محاسبه TOTAL QC SCORE", expanded=False):
            person_row = stats[stats["آیدی"].astype(str) == str(sel_id)].iloc[0]
            qc_details = compute_total_qc_score_details(person_row)
            
            st.markdown("### خلاصه")
            summary = qc_details.pop("_summary")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("جمع امتیازها", f"{summary['جمع امتیازها']} / {summary['حداکثر امتیاز']}")
            with col2:
                st.metric("نمره نهایی", f"{summary['نمره نهایی']}%")
            with col3:
                total_score = float(summary['نمره نهایی'])
                st.metric("درصد از حداکثر", f"{(total_score/100)*100:.1f}%")
            
            st.markdown("---")
            st.markdown("### جزئیات هر متریک")
            
            # دسته‌بندی متریک‌ها
            categories = {
                "متریک‌های رفتاری": [
                    "خوداظهاری اشتباه",
                    "ساعات طلایی",
                    "میانگین ساعات کاری",
                    "مشارکت با تاخیر",
                    "مشارکت‌های ناقص"
                ],
                "QC ناظران میدانی": [
                    "banner_daily",
                    "کیفیت بنر (bannerQuality)",
                    "کیفیت حضور میدانی (presenceQuality)"
                ],
                "QC راننده‌ها": [
                    "نمره حضور توسط راننده‌ها",
                    "نمره آموزش توسط راننده‌ها"
                ],
                "سایر متریک‌ها": [
                    "وعده غیرواقعی",
                    "تطبیق (نقص مدارک)",
                    "فعال‌سازی"
                ]
            }
            
            for cat_name, metrics in categories.items():
                st.markdown(f"#### {cat_name}")
                cat_data = []
                for metric in metrics:
                    if metric in qc_details:
                        detail = qc_details[metric]
                        cat_data.append({
                            "متریک": metric,
                            "مقدار": detail["مقدار"],
                            "امتیاز": detail["امتیاز"],
                            "حداکثر": detail["حداکثر"]
                        })
                
                if cat_data:
                    df_cat = pd.DataFrame(cat_data)
                    st.dataframe(df_cat, use_container_width=True, hide_index=True)
                st.markdown("")

        # === Weekly 4-week trend chart (Altair) ===
        spacer(12)
        st.markdown(
        """
        <div style='background:linear-gradient(90deg,#2563eb,#06b6d4);
            color:white;padding:10px 16px;border-radius:10px;
            font-weight:700;font-size:16px;margin-bottom:8px;'>
        📈 ترند ۴ هفته‌ای متریک‌ها (درصد نسبت به تخصیص)
        </div>
        """,
        unsafe_allow_html=True
        )
        spacer(6)

        if not mp.get("week"):
            st.error("ستون «week» پیدا نشد. اگر نامش متفاوت است، در find_column_mapping آن را مپ کن (مثلاً 'هفته').")
        else:
            # ❗️ترند را از df_raw بگیر (نه df7) تا ۴ هفتهٔ آخر واقعی محاسبه شود
            df_trend = compute_weekly_rates_for_agent(df_raw, mp, sel_id)

            if not _ALT_OK:
                st.warning("Altair نصب نیست. برای فعال شدن نمودار اجرا کن: pip install altair vega_datasets")

            if df_trend.empty:
                st.info("برای این آیدی دیتای معتبر ۴ هفته‌ای پیدا نشد (week/assigned را چک کن).")
            else:
                df_trend["week"] = pd.to_numeric(df_trend["week"], errors="coerce")
                df_trend = df_trend.dropna(subset=["week","value"])
                if not df_trend.empty and _ALT_OK:
                    try:
                        color_scale = alt.Scale(
                            domain=[
                                "روزهای بدون مشارکت",
                                "مشارکت‌های ناقص",
                                "مشارکت با تاخیر",
                                "ساعات طلایی شیفت",
                                "مشارکت خارج از محدوده مجاز",
                                "خوداظهاری اشتباه",
                            ],
                            range=["#ef4444","#f97316","#facc15","#22c55e","#3b82f6","#a855f7"]
                        )
                        chart = (
                            alt.Chart(df_trend)
                               .mark_line(point=True, strokeWidth=2)
                               .encode(
                                   x=alt.X("week:O", sort="descending", title="هفته"),
                                   y=alt.Y("value:Q", title="درصد (%)"),
                                   color=alt.Color("metric:N", title="متریک", scale=color_scale,
                                                   legend=alt.Legend(orient="bottom")),
                                   tooltip=[
                                       alt.Tooltip("metric:N", title="متریک"),
                                       alt.Tooltip("week:O", title="هفته"),
                                       alt.Tooltip("value:Q", title="درصد", format=".1f")
                                   ]
                               )
                               .properties(height=320)
                        )
                        text = (
                            alt.Chart(df_trend)
                               .mark_text(align='center', dy=-10, fontSize=11)
                               .encode(
                                   x=alt.X('week:O', sort="descending"),
                                   y='value:Q',
                                   text=alt.Text('value:Q', format='.0f'),
                                   color=alt.Color('metric:N', legend=None)
                               )
                        )
                        mean_line = alt.Chart(df_trend).mark_rule(color='gray', strokeDash=[4,4]).encode(
                            y='mean(value):Q'
                        )
                        chart = (chart + text + mean_line).configure_view(strokeWidth=0, fill="#fafafa").configure_axis(
                            grid=True, gridColor="#e5e7eb", labelFontSize=12, titleFontSize=13
                        )

                        st.altair_chart(chart, use_container_width=True)
                    except Exception as e:
                        st.warning(f"Altair خطا داد: {e}")

            with st.expander("جدول مبنای نمودار", expanded=False):
                st.dataframe(df_trend, use_container_width=True)

        spacer(8)
        st.markdown("---")



        # --- Batch download by city (collapsed) ---
        with st.expander("دانلود کارت‌ها بر اساس شهر و سرپرست (اختیاری)", expanded=False):
            cities = ["همه"] + sorted(stats["شهر"].dropna().unique())
            sel_city = st.selectbox("شهر:", cities, index=0)

            # فهرست سرپرست‌ها بر اساس شهر انتخاب‌شده
            sup_src = stats if sel_city == "همه" else stats[stats["شهر"] == sel_city]
            supervisors = ["همه"] + sorted(sup_src["سرپرست"].dropna().unique())
            sel_sup = st.selectbox("سرپرست:", supervisors, index=0)

            batch_size = st.number_input("تعداد در هر دسته", min_value=1, max_value=200, value=10, step=1)

            # اعمال فیلتر شهر/سرپرست
            subset = stats.copy()
            if sel_city != "همه":
                subset = subset[subset["شهر"] == sel_city]
            if sel_sup != "همه":
                subset = subset[subset["سرپرست"] == sel_sup]

            total_n = len(subset)
            if total_n == 0:
                st.info("برای این فیلتر کارت فعالی وجود ندارد.")
            else:
                total_batches = math.ceil(total_n / batch_size)
                batch_no = st.number_input("شمارهٔ دسته", min_value=1, max_value=total_batches, value=1, step=1)
                start_idx = (batch_no - 1) * batch_size
                end_idx = min(start_idx + batch_size, total_n)
                st.caption(f"نمایش/دانلود دستهٔ {batch_no} از {total_batches} (ردیف‌های {start_idx+1} تا {end_idx} از {total_n})")

                current_slice = subset.iloc[start_idx:end_idx]
                batch_body = ""
                city_ranks_df = compute_city_scores_ranks(stats)
                team_ranks_df = compute_team_scores_ranks(stats)
                for _, r in current_slice.iterrows():
                    rid = str(r["آیدی"])

                    # ✅ محاسبه تعداد راننده‌های فعال برای 4 هفته اخیر (همانند کارت تکی)
                    # استفاده از df_raw برای دسترسی به همه داده‌ها
                    _id_series_raw_batch = df_raw[mp["id"]].astype(str).str.strip()
                    df_agent_batch_4weeks = df_raw[_id_series_raw_batch == rid.strip()].copy()

                    if not df_agent_batch_4weeks.empty and mp.get("week"):
                        # پیدا کردن 4 هفته اخیر بر اساس ستون week
                        df_agent_batch_4weeks["_week_num"] = to_numeric_clean(df_agent_batch_4weeks[mp["week"]])
                        df_agent_batch_4weeks = df_agent_batch_4weeks.dropna(subset=["_week_num"])
                        
                        if not df_agent_batch_4weeks.empty:
                            weeks_sorted_batch = sorted(df_agent_batch_4weeks["_week_num"].unique())
                            if weeks_sorted_batch:
                                last4_weeks_batch = weeks_sorted_batch[-4:]
                                df_agent_batch_4weeks = df_agent_batch_4weeks[df_agent_batch_4weeks["_week_num"].isin(last4_weeks_batch)]
                                
                                # محاسبه تعداد راننده‌های فعال و درصد برای 4 هفته اخیر
                                if mp.get("active_drivers") and mp["active_drivers"] in df_agent_batch_4weeks.columns:
                                    active_sum = to_numeric_clean(df_agent_batch_4weeks[mp["active_drivers"]]).sum()
                                    # مقدار محاسبه‌شده را مستقیماً روی خود ر اعمال می‌کنیم تا در کارت نمایش داده شود
                                    r["تعداد راننده‌های فعال"] = active_sum

                                    if mp.get("acq") and mp["acq"] in df_agent_batch_4weeks.columns:
                                        acq_sum = to_numeric_clean(df_agent_batch_4weeks[mp["acq"]]).sum()
                                        if acq_sum > 0:
                                            r["درصد راننده‌های فعال"] = (active_sum / acq_sum) * 100
                                        else:
                                            r["درصد راننده‌های فعال"] = None
                                    else:
                                        r["درصد راننده‌های فعال"] = None
                                else:
                                    r["تعداد راننده‌های فعال"] = None
                                    r["درصد راننده‌های فعال"] = None
                                
                                # ✅ محاسبه تعداد نقص مدارک و مقادیر هر ستون برای ۴ هفته اخیر
                                # مشابه منطق کارت تکی: جمع مقدار هر ستون (صرف‌نظر از اینکه 1 یا 2 باشد) و محاسبه مجموع کلی
                                doc_issue_cols_present_batch = [col for col in DOC_ISSUE_COLS if col in df_agent_batch_4weeks.columns]
                                if doc_issue_cols_present_batch:
                                    total_doc_issues_batch = 0.0
                                    for col in doc_issue_cols_present_batch:
                                        try:
                                            col_numeric = to_numeric_clean(df_agent_batch_4weeks[col])
                                            col_sum = col_numeric.sum()
                                        except Exception:
                                            col_sum = 0
                                        # اگر NaN بود به‌عنوان صفر حساب کنیم
                                        if pd.isna(col_sum):
                                            col_sum = 0
                                        # مقدار جمع‌شده را روی ر ذخیره کنیم تا در کارت نمایش داده شود
                                        r[col] = col_sum
                                        # فقط مقادیر مثبت در مجموع لحاظ شوند
                                        try:
                                            if float(col_sum) > 0:
                                                total_doc_issues_batch += float(col_sum)
                                        except Exception:
                                            pass
                                    # اگر مجموع کلی مثبت است، همان را ذخیره کن، در غیر این صورت صفر
                                    r["تعداد نقص مدارک"] = total_doc_issues_batch if total_doc_issues_batch > 0 else 0
                                else:
                                    r["تعداد نقص مدارک"] = None
                                
                                # ✅ محاسبه درصد نقص مدارک بر اساس نسبت تعداد نقص مدارک به مجموع جذب این ۴ هفته
                                try:
                                    # مجموع جذب این ۴ هفته برای محاسبه درصد
                                    acq_sum_for_doc = None
                                    if mp.get("acq") and mp["acq"] in df_agent_batch_4weeks.columns:
                                        acq_sum_for_doc = to_numeric_clean(df_agent_batch_4weeks[mp["acq"]]).sum()
                                    total_docs_batch = r.get("تعداد نقص مدارک")
                                    if acq_sum_for_doc is not None and acq_sum_for_doc > 0 and total_docs_batch is not None:
                                        r["درصد نقص مدارک"] = (float(total_docs_batch) / float(acq_sum_for_doc)) * 100
                                    else:
                                        # اگر جذب صفر یا None باشد
                                        if total_docs_batch is not None and float(total_docs_batch) == 0:
                                            r["درصد نقص مدارک"] = 0
                                        else:
                                            r["درصد نقص مدارک"] = None
                                except Exception:
                                    r["درصد نقص مدارک"] = None
                            else:
                                r["تعداد راننده‌های فعال"] = None
                                r["درصد راننده‌های فعال"] = None
                                r["تعداد نقص مدارک"] = None
                                r["درصد نقص مدارک"] = None
                        else:
                            r["تعداد راننده‌های فعال"] = None
                            r["درصد راننده‌های فعال"] = None
                            r["تعداد نقص مدارک"] = None
                            r["درصد نقص مدارک"] = None
                    else:
                        r["تعداد راننده‌های فعال"] = None
                        r["درصد راننده‌های فعال"] = None
                        r["تعداد نقص مدارک"] = None
                        r["درصد نقص مدارک"] = None
                    
                    agg = build_acq_trend_last4weeks(df_raw, mp, rid)
                    chart_src = build_acq_trend_svg_from_df(agg, width=420, height=160, point_fs=14)
                    recent3   = acq_last_k_weeks_sum(df_raw, mp, rid, k=4)
                    last_week = acq_last_week_sum(df_raw, mp, rid)

                    rr = city_ranks_df[city_ranks_df["آیدی"].astype(str) == rid]
                    rt = f'{int(rr.iloc[0]["CityRank"])} از {int(rr.iloc[0]["CityCount"])}' if not rr.empty else None

                    team_rr = team_ranks_df[team_ranks_df["آیدی"].astype(str) == rid]
                    team_rt = f'{int(team_rr.iloc[0]["TeamRank"])} از {int(team_rr.iloc[0]["TeamCount"])}' if not team_rr.empty else None

                    batch_body += build_card_html(
                        r.to_dict(), from_j_ref, to_j_ref,
                        logo_src=logo_b64,
                        rank_text=rt, team_rank_text=team_rt,
                        acq_chart_src=chart_src,
                        acq_recent3_total=recent3,
                        acq_last_week=last_week,
                    )

                batch_html = build_agent_date_block_shamsi(from_j_ref, to_j_ref, logo_src=logo_b64) + batch_body

                # نام فایل خروجی بر اساس شهر/سرپرست
                file_city = (sel_city if sel_city != "همه" else "all_cities").replace(" ", "_")
                file_sup  = (sel_sup  if sel_sup  != "همه" else "all_supervisors").replace(" ", "_")

                html_download_button(
                    f"cards_{file_city}_{file_sup}_batch{batch_no}_of_{total_batches}.html",
                    batch_html,
                    "دانلود HTML این دسته"
                )

                # --- لینک دانلود همهٔ دسته‌ها (هر دسته یک فایل) ---
                with st.expander("لینک دانلود همهٔ دسته‌ها (هر دسته یک فایل)", expanded=False):
                    for b in range(1, total_batches+1):
                        s = (b-1)*batch_size
                        e = min(s+batch_size, total_n)
                        sl = subset.iloc[s:e]

                        body = ""
                        for _, r in sl.iterrows():
                            rid = str(r["آیدی"])

                            # رتبه شهر
                            rr = city_ranks_df[city_ranks_df["آیدی"].astype(str) == rid]
                            rt = f'{int(rr.iloc[0]["CityRank"])} از {int(rr.iloc[0]["CityCount"])}' if not rr.empty else None

                            # رتبه تیم
                            team_rr = team_ranks_df[team_ranks_df["آیدی"].astype(str) == rid]
                            team_rt = f'{int(team_rr.iloc[0]["TeamRank"])} از {int(team_rr.iloc[0]["TeamCount"])}' if not team_rr.empty else None

                            # ✅ نمودار و جمع سه‌هفته‌ای
                            agg = build_acq_trend_last4weeks(df_raw, mp, rid)
                            chart_src = build_acq_trend_svg_from_df(agg, width=420, height=160, point_fs=14)
                            recent3   = acq_last_k_weeks_sum(df_raw, mp, rid, k=4)
                            last_week = acq_last_week_sum(df_raw, mp, rid)

                            # کارت با همه‌ی پارامترهای لازم
                            body += build_card_html(
                                r.to_dict(), from_j_ref, to_j_ref,
                                logo_src=logo_b64,
                                rank_text=rt, team_rank_text=team_rt,
                                acq_chart_src=chart_src,
                                acq_recent3_total=recent3,
                                acq_last_week=last_week,
                            )

                        out_html = build_agent_date_block_shamsi(from_j_ref, to_j_ref, logo_src=logo_b64) + body
                        fname = f"cards_{file_city}_{file_sup}_batch{b}_of_{total_batches}.html"
                        b64 = base64.b64encode(out_html.encode("utf-8")).decode()
                        st.markdown(f'• <a download="{fname}" href="data:text/html;base64,{b64}">{fname}</a>', unsafe_allow_html=True)

    # ======== MODE 2: Supervisor ========
    else:
        sup_day_all = compute_supervisor_daily(df7, mp)
        ov_total_all = compute_supervisor_overview_total(sup_day_all)
        
        # ✅ میانگین TOTAL QC SCORE و QC ناظران / راننده ها برای هر سرپرست در همین بازه
        stats_sup = compute_weekly_stats(df7, mp)

        # --- میانگین TOTAL QC SCORE ---
        if "TOTAL QC SCORE" in stats_sup.columns:
            qc_sup = (
                stats_sup
                .groupby(["سرپرست", "شهر"], dropna=False)["TOTAL QC SCORE"]
                .mean()
                .reset_index()
                .rename(columns={"TOTAL QC SCORE": "میانگین TOTAL QC SCORE"})
            )
            ov_total_all = ov_total_all.merge(qc_sup, on=["سرپرست","شهر"], how="left")

        # --- میانگین QC ناظران میدانی (درصد) ---
        if "نمره QC ناظران میدانی" in stats_sup.columns:
            qc_field_sup = (
                stats_sup
                .groupby(["سرپرست","شهر"], dropna=False)["نمره QC ناظران میدانی"]
                .mean()
                .reset_index()
                .rename(columns={"نمره QC ناظران میدانی": "میانگین QC ناظران میدانی"})
            )
            ov_total_all = ov_total_all.merge(qc_field_sup, on=["سرپرست","شهر"], how="left")

        # --- میانگین QC راننده ها (لیبل → عدد → میانگین) ---
        if "نمره QC راننده ها" in stats_sup.columns:
            mapping = {"عالی":4, "خوب":3, "متوسط":2, "ضعیف":1}
            tmp = stats_sup.copy()
            tmp["_qc_driver_num"] = tmp["نمره QC راننده ها"].map(mapping)

            qc_driver_sup = (
                tmp
                .groupby(["سرپرست","شهر"], dropna=False)["_qc_driver_num"]
                .mean()
                .reset_index()
                .rename(columns={"_qc_driver_num": "میانگین QC راننده ها"})
            )
            ov_total_all = ov_total_all.merge(qc_driver_sup, on=["سرپرست","شهر"], how="left")

        st.subheader("فیلتر شهر و سرپرست")

        cities = ["همه"] + sorted(sup_day_all["شهر"].dropna().unique())
        sel_city = st.selectbox("شهر:", cities, index=0)
        sup_list = sup_day_all["سرپرست"] if sel_city == "همه" else sup_day_all[sup_day_all["شهر"] == sel_city]["سرپرست"]
        supervisors = sorted(sup_list.dropna().unique())
        sel_sup = st.selectbox("سرپرست:", supervisors)

        day_tbl = sup_day_all[sup_day_all["سرپرست"] == sel_sup]
        ov_tbl = ov_total_all[ov_total_all["سرپرست"] == sel_sup]
        if sel_city != "همه":
            day_tbl = day_tbl[day_tbl["شهر"] == sel_city]
            ov_tbl = ov_tbl[ov_tbl["شهر"] == sel_city]

        if not day_tbl.empty:
            _from_g = pd.to_datetime(day_tbl["تاریخ"]).min().date()
            _to_g   = pd.to_datetime(day_tbl["تاریخ"]).max().date()
        else:
            _to_g = anchor
            _from_g = anchor - timedelta(days=6)

        date_from_disp = to_jalali_words(_from_g, persian_digits=True)
        date_to_disp   = to_jalali_words(_to_g,   persian_digits=True)

        city_disp = sel_city if sel_city != "همه" else "، ".join(sorted(day_tbl["شهر"].dropna().unique()))
        html = build_supervisor_html_persian(
            sel_sup,
            city_disp,
            from_j_ref,          # ✅ تاریخ شروع مرجع (week آخر)
            to_j_ref,            # ✅ تاریخ پایان مرجع (week آخر)
            overview_df=ov_tbl,
            daily_df=day_tbl,
            logo_src=logo_b64
        )
        st.components.v1.html(html, height=900, scrolling=True)
        fname = f"supervisor_{sel_sup}_{city_disp}.html".replace(" ", "_")
        b64 = base64.b64encode(html.encode("utf-8")).decode()
        st.markdown(f'<a download="{fname}" href="data:text/html;base64,{b64}">دانلود HTML گزارش سرپرست</a>', unsafe_allow_html=True)
