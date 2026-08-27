import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import date

st.set_page_config(page_title="جودة بيانات الموظفين", page_icon="📊", layout="wide")

st.markdown("""
<style>
html, body, [class*="css"] { direction: rtl; text-align: right; }
[data-testid="stMetricValue"] { direction:ltr; }
</style>
""", unsafe_allow_html=True)

st.title("📊 نظام تدقيق جودة بيانات الموظفين")
st.caption("ارفع ملف Excel أو CSV وسيتم تطبيق قواعد جودة البيانات تلقائياً دون تعديل الملف الأصلي.")

TODAY = pd.Timestamp.today().normalize()

UAE_EMIRATES = {
    "أبوظبي", "دبي", "الشارقة", "عجمان", "أم القيوين", "رأس الخيمة", "الفجيرة"
}

GCC_KEYWORDS = {
    "الإمارات", "الامارات", "إماراتي", "اماراتي", "إماراتية", "اماراتية",
    "السعودية", "سعودي", "سعودية",
    "الكويت", "كويتي", "كويتية",
    "البحرين", "بحريني", "بحرينية",
    "قطر", "قطري", "قطرية",
    "عمان", "عُمان", "عماني", "عمانية"
}

UNIQUE_COLS = ["الرقم الوظيفي", "اسم الموظف", "الاسم بالانجليزي", "رقم الهوية", "رقم الجواز"]

DATE_COLS = [
    "تاريخ التعيين", "تاريخ الميلاد", "تاريخ انتهاء الدارسة", "تاريخ انتهاء الدراسة",
    "تاريخ التصديق", "تاريخ إنتهاء الهوية", "تاريخ انتهاء الهوية",
    "تاريخ اصدار الجواز", "تاريخ إصدار الجواز", "تاريخ إنتهاء الجواز", "تاريخ انتهاء الجواز",
    "تاريخ اصدار اللإقامة", "تاريخ إصدار الإقامة", "تاريخ انتهاء اللإقامة", "تاريخ انتهاء الإقامة"
]

RESIDENCY_FIELDS = [
    "رقم الأقامة", "رقم الإقامة", "الكفيل",
    "تاريخ اصدار اللإقامة", "تاريخ إصدار الإقامة",
    "تاريخ انتهاء اللإقامة", "تاريخ انتهاء الإقامة"
]

def clean_text(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def is_blank(v):
    return pd.isna(v) or clean_text(v).lower() in {"", "nan", "none", "null", "nat"}

def normalize_name(v):
    s = clean_text(v)
    s = re.sub(r"\s+", " ", s)
    return s.casefold()

def first_name(v):
    s = clean_text(v)
    if not s:
        return ""
    return re.split(r"\s+", s)[0].strip().casefold()

def parse_date(v):
    if is_blank(v):
        return pd.NaT
    try:
        return pd.to_datetime(v, errors="coerce", dayfirst=True)
    except Exception:
        return pd.NaT

def years_between(start, end):
    if pd.isna(start) or pd.isna(end):
        return np.nan
    return end.year - start.year - ((end.month, end.day) < (start.month, start.day))

def add_issue(issues, df, idx, field, category, severity, value, message):
    issues.append({
        "رقم الصف": int(idx) + 2,
        "الرقم الوظيفي": clean_text(df.at[idx, "الرقم الوظيفي"]) if "الرقم الوظيفي" in df.columns else "",
        "اسم الموظف": clean_text(df.at[idx, "اسم الموظف"]) if "اسم الموظف" in df.columns else "",
        "الحقل": field,
        "نوع الفحص": category,
        "الخطورة": severity,
        "القيمة الحالية": clean_text(value),
        "الملاحظة": message
    })

def find_col(df, options):
    for c in options:
        if c in df.columns:
            return c
    return None

def is_gcc(nationality):
    n = clean_text(nationality).replace("ـ", "")
    return any(k in n for k in GCC_KEYWORDS)

def audit(df, gender_min_count=5, gender_threshold=90):
    issues = []
    checked_rules = []

    # 1) Uniqueness
    for col in UNIQUE_COLS:
        if col not in df.columns:
            continue
        checked_rules.append(f"عدم التكرار: {col}")
        normalized = df[col].map(normalize_name)
        valid = ~df[col].map(is_blank)
        dup = normalized.duplicated(keep=False) & valid
        for idx in df.index[dup]:
            add_issue(issues, df, idx, col, "Uniqueness", "عالية", df.at[idx, col],
                      f"القيمة مكررة في {int((normalized == normalized.at[idx]).sum())} سجلات.")

    # 2) Appointment date vs service duration
    if "تاريخ التعيين" in df.columns and "مدة الخدة" in df.columns:
        checked_rules.append("تطابق تاريخ التعيين مع مدة الخدمة")
        for idx, row in df.iterrows():
            d = parse_date(row["تاريخ التعيين"])
            raw = row["مدة الخدة"]
            if pd.isna(d) or is_blank(raw):
                continue
            try:
                stated = float(str(raw).replace(",", ".").strip())
                calculated = years_between(d, TODAY)
                if abs(stated - calculated) > 1:
                    add_issue(issues, df, idx, "مدة الخدة", "Consistency", "متوسطة", raw,
                              f"مدة الخدمة المسجلة {stated:g}، بينما المحسوبة من تاريخ التعيين حوالي {calculated} سنة.")
            except:
                add_issue(issues, df, idx, "مدة الخدة", "Validity", "متوسطة", raw,
                          "مدة الخدمة ليست قيمة رقمية قابلة للمقارنة.")

    # 3) DOB vs age
    if "تاريخ الميلاد" in df.columns and "العمر" in df.columns:
        checked_rules.append("تطابق تاريخ الميلاد مع العمر")
        for idx, row in df.iterrows():
            d = parse_date(row["تاريخ الميلاد"])
            raw = row["العمر"]
            if pd.isna(d) or is_blank(raw):
                continue
            try:
                stated = int(float(str(raw).strip()))
                calculated = years_between(d, TODAY)
                if stated != calculated:
                    add_issue(issues, df, idx, "العمر", "Consistency", "متوسطة", raw,
                              f"العمر المسجل {stated}، بينما العمر المحسوب من تاريخ الميلاد {calculated}.")
            except:
                add_issue(issues, df, idx, "العمر", "Validity", "متوسطة", raw, "العمر ليس رقماً صحيحاً.")

    # 4) Gender anomaly from first-name majority
    if "اسم الموظف" in df.columns and "الجنس" in df.columns:
        checked_rules.append("تحليل شذوذ الجنس من تكرار الاسم الأول")
        temp = pd.DataFrame({
            "first": df["اسم الموظف"].map(first_name),
            "gender": df["الجنس"].map(lambda x: clean_text(x).casefold())
        }, index=df.index)
        temp = temp[(temp["first"] != "") & (temp["gender"] != "")]
        for fname, grp in temp.groupby("first"):
            total = len(grp)
            if total < gender_min_count:
                continue
            counts = grp["gender"].value_counts()
            if counts.empty:
                continue
            dominant = counts.index[0]
            pct = counts.iloc[0] / total * 100
            if pct >= gender_threshold:
                for idx in grp.index[grp["gender"] != dominant]:
                    add_issue(issues, df, idx, "الجنس", "Anomaly Detection", "متوسطة", df.at[idx, "الجنس"],
                              f"الاسم الأول «{fname}» ظهر {total} مرة، و{pct:.1f}% منه مسجل كـ «{dominant}». هذه القيمة مخالفة للنمط الغالب وتحتاج مراجعة.")

    # 5) Qualification relevance language/values
    qcol = "هل المؤهل متوافق للوظيفة"
    if qcol in df.columns:
        checked_rules.append("توحيد حقل توافق المؤهل")
        ar_allowed = {"مطابق للوظيفة", "غير مطابق للوظيفة", "مطابق للوظيفه", "غير مطابق للوظيفه"}
        en_allowed = {"relevant", "irrelevant"}
        for idx, v in df[qcol].items():
            if is_blank(v):
                add_issue(issues, df, idx, qcol, "Completeness", "منخفضة", v, "القيمة فارغة / NULL.")
                continue
            s = clean_text(v)
            has_ar = bool(re.search(r"[\u0600-\u06FF]", s))
            has_en = bool(re.search(r"[A-Za-z]", s))
            if has_ar and has_en:
                add_issue(issues, df, idx, qcol, "Consistency", "متوسطة", v,
                          "القيمة تحتوي العربية والإنجليزية معاً.")
            elif s.casefold() not in en_allowed and s not in ar_allowed:
                add_issue(issues, df, idx, qcol, "Validity", "متوسطة", v,
                          "القيمة خارج القيم المتوقعة: Relevant / Irrelevant / مطابق للوظيفة / غير مطابق للوظيفة.")

    # 6) All date fields must be valid dates
    for col in DATE_COLS:
        if col not in df.columns:
            continue
        checked_rules.append(f"صلاحية التاريخ: {col}")
        for idx, v in df[col].items():
            if not is_blank(v) and pd.isna(parse_date(v)):
                add_issue(issues, df, idx, col, "Data Type", "عالية", v,
                          "القيمة غير قابلة للتعريف كتاريخ صحيح.")

    # 7) Emirates
    emirate_col = find_col(df, ["العنوان امارة", "العنوان إمارة"])
    if emirate_col:
        checked_rules.append("توحيد أسماء الإمارات السبع بالعربية")
        for idx, v in df[emirate_col].items():
            if not is_blank(v) and clean_text(v) not in UAE_EMIRATES:
                add_issue(issues, df, idx, emirate_col, "Validity", "متوسطة", v,
                          "يجب أن تكون القيمة أحد أسماء الإمارات السبع باللغة العربية وبالصيغة الموحدة.")

    # 8) Phone +971
    phone_col = "رقم الهاتف" if "رقم الهاتف" in df.columns else None
    if phone_col:
        checked_rules.append("نمط رقم الهاتف +971")
        pattern = re.compile(r"^\+971\d{7,9}$")
        for idx, v in df[phone_col].items():
            if is_blank(v):
                continue
            compact = re.sub(r"[\s\-()]", "", clean_text(v))
            if not pattern.fullmatch(compact):
                add_issue(issues, df, idx, phone_col, "Format", "متوسطة", v,
                          "رقم الهاتف لا يتبع النمط الموحد الذي يبدأ بـ +971.")

    # 9) Emirates ID format
    id_col = "رقم الهوية" if "رقم الهوية" in df.columns else None
    if id_col:
        checked_rules.append("نمط رقم الهوية 784-YYYY-NNNNNNN-C")
        pattern = re.compile(r"^784-\d{4}-\d{7}-\d$")
        for idx, v in df[id_col].items():
            if is_blank(v):
                continue
            s = clean_text(v)
            if not pattern.fullmatch(s):
                add_issue(issues, df, idx, id_col, "Format", "عالية", v,
                          "رقم الهوية يجب أن يبدأ بـ 784- ويتبع مثال: 784-1927-0591531-8.")

    # 10) Residency expiry + conditional residency completeness
    expiry_col = find_col(df, ["تاريخ انتهاء اللإقامة", "تاريخ انتهاء الإقامة"])
    issue_col = find_col(df, ["تاريخ اصدار اللإقامة", "تاريخ إصدار الإقامة"])
    nationality_col = "الجنسية" if "الجنسية" in df.columns else None

    if expiry_col:
        checked_rules.append("تنبيه انتهاء الإقامة")
        for idx, v in df[expiry_col].items():
            d = parse_date(v)
            if not is_blank(v) and not pd.isna(d) and d.normalize() < TODAY:
                add_issue(issues, df, idx, expiry_col, "Expiry", "عالية", v,
                          f"الإقامة منتهية مقارنة بتاريخ اليوم {TODAY.date()}.")

    if issue_col and expiry_col:
        checked_rules.append("تاريخ إصدار الإقامة قبل تاريخ انتهائها")
        for idx, row in df.iterrows():
            a, b = parse_date(row[issue_col]), parse_date(row[expiry_col])
            if not pd.isna(a) and not pd.isna(b) and a > b:
                add_issue(issues, df, idx, expiry_col, "Consistency", "عالية", row[expiry_col],
                          "تاريخ انتهاء الإقامة أقدم من تاريخ إصدار الإقامة.")

    if nationality_col:
        checked_rules.append("بيانات الإقامة مطلوبة لغير الإماراتيين ومواطني دول مجلس التعاون")
        existing_res_fields = [c for c in RESIDENCY_FIELDS if c in df.columns]
        for idx, nationality in df[nationality_col].items():
            if is_blank(nationality) or is_gcc(nationality):
                continue
            for col in existing_res_fields:
                if is_blank(df.at[idx, col]):
                    add_issue(issues, df, idx, col, "Conditional Completeness", "عالية", df.at[idx, col],
                              f"بيانات الإقامة مطلوبة لأن الجنسية «{clean_text(nationality)}» ليست من دول مجلس التعاون الخليجي.")

    # 11) Email must be valid and end with ae
    email_col = "البريد الالكتروني" if "البريد الالكتروني" in df.columns else None
    if email_col:
        checked_rules.append("اتساق البريد الإلكتروني ونهايته ae")
        email_pattern = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", re.I)
        for idx, v in df[email_col].items():
            if is_blank(v):
                continue
            s = clean_text(v)
            if not email_pattern.fullmatch(s):
                add_issue(issues, df, idx, email_col, "Format", "عالية", v,
                          "صيغة البريد الإلكتروني غير صحيحة.")
            elif not s.casefold().endswith("ae"):
                add_issue(issues, df, idx, email_col, "Consistency", "متوسطة", v,
                          "البريد الإلكتروني لا ينتهي بالحرفين ae.")

    result = pd.DataFrame(issues)
    return result, sorted(set(checked_rules))

def to_excel(original, issues):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        issues.to_excel(writer, index=False, sheet_name="الأخطاء")
        original.to_excel(writer, index=False, sheet_name="البيانات الأصلية")
        if not issues.empty:
            summary = (issues.groupby(["الحقل", "نوع الفحص", "الخطورة"])
                       .size().reset_index(name="عدد الأخطاء")
                       .sort_values("عدد الأخطاء", ascending=False))
            summary.to_excel(writer, index=False, sheet_name="ملخص")
    return out.getvalue()

with st.sidebar:
    st.header("⚙️ إعدادات التحليل")
    gender_min = st.number_input("الحد الأدنى لتكرار الاسم", min_value=2, max_value=100, value=5)
    gender_pct = st.slider("نسبة الأغلبية لتنبيه الجنس", 50, 100, 90, 1)
    st.info("مثال: إذا تكرر اسم فاطمة 100 مرة، 99 أنثى و1 ذكر، سيتم تنبيه السجل المخالف.")

uploaded = st.file_uploader("ارفع ملف بيانات الموظفين", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.info("⬆️ ارفعي ملف Excel أو CSV لبدء التدقيق.")
    st.stop()

try:
    filename = uploaded.name.lower()
    file_bytes = uploaded.getvalue()

    if filename.endswith(".csv"):
        last_error = None
        for enc in ["utf-8-sig", "utf-8", "cp1256", "latin1"]:
            try:
                df = pd.read_csv(BytesIO(file_bytes), encoding=enc)
                last_error = None
                break
            except Exception as err:
                last_error = err
        if last_error:
            raise last_error
    elif filename.endswith(".xlsx"):
        df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
    elif filename.endswith(".xls"):
        try:
            df = pd.read_excel(BytesIO(file_bytes), engine="xlrd")
        except Exception:
            try:
                tables = pd.read_html(BytesIO(file_bytes))
                df = tables[0]
            except Exception as err:
                raise ValueError("الملف .xls غير صالح كملف Excel قديم. افتحيه في Excel ثم Save As → Excel Workbook (*.xlsx).") from err
    else:
        raise ValueError("نوع الملف غير مدعوم. استخدمي xlsx أو xls أو csv.")
except Exception as e:
    st.error(f"تعذر قراءة الملف: {e}")
    st.info("إذا كان الملف صادر من Oracle/ERP أو تقرير قديم: افتحيه في Excel ثم Save As → Excel Workbook (*.xlsx) وبعدها ارفعي النسخة الجديدة.")
    st.stop()

df.columns = [str(c).strip() for c in df.columns]
issues, checked_rules = audit(df, int(gender_min), int(gender_pct))

st.success(f"تمت قراءة {len(df):,} سجل و {len(df.columns):,} حقل.")

total_cells = max(len(df) * len(df.columns), 1)
error_rows = issues["رقم الصف"].nunique() if not issues.empty else 0
clean_rows = max(len(df) - error_rows, 0)
quality = max(0, 100 * (1 - len(issues) / total_cells))

c1, c2, c3, c4 = st.columns(4)
c1.metric("نسبة الجودة التقريبية", f"{quality:.1f}%")
c2.metric("إجمالي الملاحظات", f"{len(issues):,}")
c3.metric("سجلات بها ملاحظات", f"{error_rows:,}")
c4.metric("سجلات بلا ملاحظات", f"{clean_rows:,}")

tabs = st.tabs(["📋 الأخطاء", "📊 الملخص", "🧠 شذوذ الجنس", "✅ القواعد المنفذة", "🗂️ البيانات"])

with tabs[0]:
    if issues.empty:
        st.success("لم يتم العثور على مخالفات وفق القواعد الحالية.")
    else:
        f1, f2, f3 = st.columns(3)
        fields = f1.multiselect("الحقل", sorted(issues["الحقل"].dropna().unique()))
        types = f2.multiselect("نوع الفحص", sorted(issues["نوع الفحص"].dropna().unique()))
        sev = f3.multiselect("الخطورة", sorted(issues["الخطورة"].dropna().unique()))
        view = issues.copy()
        if fields: view = view[view["الحقل"].isin(fields)]
        if types: view = view[view["نوع الفحص"].isin(types)]
        if sev: view = view[view["الخطورة"].isin(sev)]
        st.dataframe(view, use_container_width=True, hide_index=True)

with tabs[1]:
    if not issues.empty:
        summary = (issues.groupby(["الحقل", "نوع الفحص"])
                   .size().reset_index(name="عدد الأخطاء")
                   .sort_values("عدد الأخطاء", ascending=False))
        st.dataframe(summary, use_container_width=True, hide_index=True)
        st.bar_chart(summary.set_index("الحقل")["عدد الأخطاء"])
    else:
        st.success("لا توجد أخطاء لعرضها.")

with tabs[2]:
    gender_issues = issues[issues["نوع الفحص"] == "Anomaly Detection"] if not issues.empty else pd.DataFrame()
    if gender_issues.empty:
        st.success("لم يتم اكتشاف حالات جنس مخالفة للنمط الغالب للأسماء.")
    else:
        st.warning(f"تم اكتشاف {len(gender_issues)} حالة تحتاج مراجعة.")
        st.dataframe(gender_issues, use_container_width=True, hide_index=True)

with tabs[3]:
    st.write(f"تم تنفيذ **{len(checked_rules)}** قاعدة/فحص متاح حسب الأعمدة الموجودة في الملف:")
    for r in checked_rules:
        st.write("✓", r)
    missing = [c for c in set(UNIQUE_COLS + ["تاريخ التعيين","مدة الخدة","تاريخ الميلاد","العمر","الجنس",
                                              "هل المؤهل متوافق للوظيفة","العنوان امارة","رقم الهاتف",
                                              "الجنسية","البريد الالكتروني"]) if c not in df.columns]
    if missing:
        st.warning("بعض الحقول المتوقعة غير موجودة بالاسم نفسه في الملف:")
        st.write("، ".join(sorted(missing)))

with tabs[4]:
    st.dataframe(df, use_container_width=True, hide_index=True)

st.divider()
excel_bytes = to_excel(df, issues)
st.download_button(
    "⬇️ تحميل تقرير جودة البيانات Excel",
    data=excel_bytes,
    file_name=f"data_quality_report_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
