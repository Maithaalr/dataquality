        st.success(
            "لم يتم اكتشاف حالات جنس "
            "مخالفة للنمط الغالب للأسماء."
        )

    else:

        st.warning(
            f"تم اكتشاف "
            f"{len(gender_issues)} "
            f"حالة تحتاج مراجعة."
        )


        st.dataframe(
            gender_issues,
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# Tab 5 - القواعد المنفذة
# =========================================================

with tabs[4]:

    st.write(
        f"تم تنفيذ **{len(checked_rules)}** "
        f"قاعدة/فحص متاح حسب الأعمدة "
        f"الموجودة في الملف:"
    )


    for r in checked_rules:

        st.write(
            "✓",
            r
        )


    st.divider()


    st.markdown(
        "### قواعد حساب الاستكمال"
    )


    st.write(
        "✓ جميع الحقول الموجودة في الملف "
        "يتم فحص نسبة استكمالها."
    )


    st.write(
        "✓ الحقول العادية مطلوبة لجميع الموظفين."
    )


    st.write(
        "✓ بيانات الإقامة مطلوبة فقط لغير "
        "مواطني دول مجلس التعاون الخليجي."
    )


    st.write(
        "✓ رقم خلاصة القيد ورقم البلدة "
        "ورقم الأسرة ب مطلوبة فقط "
        "للمواطنين الإماراتيين."
    )


    st.write(
        "✓ عدد الأبناء يعتبر مكتمل حتى "
        "إذا كانت القيمة فارغة."
    )


    st.write(
        "✓ المؤسسة التعليمية وتاريخ انتهاء "
        "الدراسة ودرجة المؤهل تُطلب فقط "
        "للمؤهلات المحددة."
    )


# =========================================================
# Tab 6 - البيانات الأصلية
# =========================================================

with tabs[5]:

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# تحميل التقرير
# =========================================================

st.divider()


excel_bytes = to_excel(
    df,
    issues,
    completeness_df,
    missing_details
)


st.download_button(

    "⬇️ تحميل تقرير جودة البيانات Excel",

    data=excel_bytes,

    file_name=(
        f"data_quality_report_"
        f"{date.today().isoformat()}.xlsx"
    ),

    mime=(
        "application/vnd.openxmlformats-"

