import streamlit as st
import pandas as pd
import os
from datetime import datetime
import re
import altair as alt
import io

# ---------- НАСТРОЙКА СТРАНИЦЫ ----------
st.set_page_config(page_title="CRM НИУ ВШЭ МИЭМ", layout="wide")

# Логотип (файл должен лежать в одной папке с app.py)
st.image("01_Logo_HSE_full_rus_Pantone.png", width=120)

# Заголовок в две строки
st.markdown("""
## CRM НИУ ВШЭ МИЭМ  
### Управление научными проектами
""")

DATA_FILE = "projects.xlsx"

# ---------- КОЛОНКИ ----------
COLUMNS = [
    "id", "supervisor", "supervisor_department", "supervisor_grnti", "supervisor_team",
    "supervisor_competencies", "supervisor_publications", "supervisor_grants", "supervisor_niokr",
    "supervisor_rid_protected", "supervisor_rid_unprotected", "supervisor_ugt", "supervisor_next_ugt",
    "project_name", "customer", "problem", "competitor", "advantage", "partner", "role_miem",
    "horizon", "funding_source", "lifecycle_stage", "sales_stage", "stage_change_reason", "stage_change_date"
]

# ---------- ФУНКЦИИ ----------
def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, dtype={"id": int})
    else:
        df = pd.DataFrame(columns=COLUMNS)
        df["id"] = df["id"].astype(int)

    for col in COLUMNS:
        if col not in df.columns:
            if col == "id":
                df[col] = 0
            elif col == "supervisor_ugt":
                df[col] = 1
            else:
                df[col] = ""
    text_cols = [c for c in COLUMNS if c not in ["id", "supervisor_ugt", "stage_change_date"]]
    for col in text_cols:
        df[col] = df[col].fillna("").astype(str)
    df["id"] = df["id"].astype(int)
    df["supervisor_ugt"] = pd.to_numeric(df["supervisor_ugt"], errors="coerce").fillna(1).astype(int)
    if "stage_change_date" in df.columns:
        df["stage_change_date"] = pd.to_datetime(df["stage_change_date"], errors="coerce")
    # Удаляем записи с пустым заказчиком или "[вручную]"
    df = df[~df["customer"].astype(str).str.contains(r"\[вручную\]", na=False, case=False)]
    df = df[df["customer"].notna() & (df["customer"].astype(str).str.strip() != "")]
    df = df.reset_index(drop=True)
    if not df.empty:
        df = df.sort_values("id").reset_index(drop=True)
    return df

def save_data(df):
    df.to_excel(DATA_FILE, index=False)

def get_next_id(df):
    return int(df["id"].max() + 1) if not df.empty else 1

# ---------- СПИСКИ ДЛЯ ВЫБОРА ----------
LIFECYCLE_STAGES = ["Планирование (НИР)", "Проектирование (ОКР)", "Разработка", "Внедрение", "Эксплуатация"]
SALES_STAGES = [
    "Квалификация", "Выявление проблем", "Формирование видения",
    "Обоснование ценности", "Проработка решения", "Презентация",
    "Переговоры и возражения", "Закрытие сделки", "Поддержка и развитие"
]
ROLES_MIEM = ["Субподрядчик", "Соисполнитель", "Лицензиар", "Сервисный центр", "Технологический аудитор",
              "Консультант", "Другая роль"]
HORIZON_LIST = ["0-3 месяца", "3-6 месяцев", "6-12 месяцев", "1-3 года", "Другой срок"]
FUNDING_SOURCES = ["Внутренний (грант ВШЭ/МИЭМ)", "Внешний институциональный (РНФ, РФФИ, Минобр)",
                   "Внешний корпоративный (компания)", "Смешанный", "Другое"]
UNPROTECTED_RID_LIST = ["Алгоритмы/методы", "ПО", "КД", "Прототип", "Ноу-хау",
                         "Публикации с тех. описанием", "Технология", "Бизнес-модель", "Другой неоф. РИД"]
DEPARTMENT_LIST = [
    "Центр квантовых метаматериалов", "Международная лаборатория физики элементарных частиц",
    "Научно-технический центр прикладной электроники", "Научно-учебная лаборатория квантовой наноэлектроники",
    "Научно-учебная лаборатория телекоммуникационных систем", "Научная лаборатория Интернета вещей и киберфизических систем",
    "Учебно-исследовательская лаборатория функциональной безопасности космических аппаратов и систем",
    "Лаборатория «Математические методы естествознания»", "Лаборатория вычислительной физики",
    "Лаборатория динамических систем и приложений", "Учебно-исследовательская лаборатория Интернет технологий и сервисов",
    "Учебная лаборатория 3Д-визуализации и компьютерной графики", "Учебная лаборатория элементов и устройств встраиваемых систем",
    "Учебная лаборатория систем автоматизированного проектирования", "Учебная лаборатория сетевых технологий",
    "Учебная лаборатория сетевых видеотехнологий", "Учебная лаборатория моделирования систем защиты информации и криптографии",
    "Учебная лаборатория общей и квантовой физики", "Учебная лаборатория квантовых технологий",
    "Учебная лаборатория метрологии и измерительных технологий", "Учебная лаборатория надежности электронных средств киберфизических систем",
    "Учебная лаборатория моделирования и проектирования электронных компонентов и устройств",
    "Учебная лаборатория микроволновой и оптоэлектронной инженерии", "Учебная лаборатория телекоммуникационных технологий и систем связи",
    "Учебная лаборатория электроники и схемотехники", "Другое"
]

# ---------- НАВИГАЦИЯ ----------
st.sidebar.title("Навигация")
page = st.sidebar.radio("Перейти", [
    "👨‍🔬 Научные руководители",
    "📋 Проекты",
    "📊 Дашборд",
    "💾 Экспорт / Импорт"
])

# ========== СТРАНИЦА "НАУЧНЫЕ РУКОВОДИТЕЛИ" ==========
if page == "👨‍🔬 Научные руководители":
    st.header("Научные руководители и их проекты")
    df = load_data()
    if df.empty:
        st.info("Нет данных. Создайте первый проект на вкладке 'Проекты'.")
    else:
        supervisors = sorted(df["supervisor"].dropna().unique())
        selected_sup = st.selectbox("Выберите научного руководителя", supervisors)
        sup_df = df[df["supervisor"] == selected_sup]
        st.subheader(f"Проекты руководителя {selected_sup}")

        with st.expander("📋 Карточка научного руководителя (задел)"):
            st.text_input("ФИО научного руководителя", value=selected_sup, disabled=True)
            current_dept = sup_df.iloc[0]["supervisor_department"] if "supervisor_department" in sup_df.columns else ""
            dept_idx = DEPARTMENT_LIST.index(current_dept) if current_dept in DEPARTMENT_LIST else 0
            department = st.selectbox("Подразделение МИЭМ (выбрать)", DEPARTMENT_LIST, index=dept_idx)
            grnti = st.text_input("Научное направление (шифр ГРНТИ)", value=sup_df.iloc[0]["supervisor_grnti"])
            team = st.text_area("Команда (количество, состав, ФИО)", value=sup_df.iloc[0]["supervisor_team"])
            comp = st.text_area("Ключевая компетенция", value=sup_df.iloc[0]["supervisor_competencies"])
            pubs = st.text_area("Публикации (до 5)", value=sup_df.iloc[0]["supervisor_publications"])
            grants = st.text_area("Гранты (фонд, период, объём)", value=sup_df.iloc[0]["supervisor_grants"])
            niokr = st.text_area("Выполненные или идущие НИОКР", value=sup_df.iloc[0]["supervisor_niokr"])
            rid_protected = st.text_area("Охраняемые РИД", value=sup_df.iloc[0]["supervisor_rid_protected"])
            current_unprot = sup_df.iloc[0]["supervisor_rid_unprotected"] if "supervisor_rid_unprotected" in sup_df.columns else ""
            unprot_idx = UNPROTECTED_RID_LIST.index(current_unprot) if current_unprot in UNPROTECTED_RID_LIST else 0
            rid_unprotected = st.selectbox("Неоформленные РИД (выбрать)", UNPROTECTED_RID_LIST, index=unprot_idx)
            ugt = st.slider("УГТ задела научной группы (выбрать)", 1, 9, value=int(sup_df.iloc[0]["supervisor_ugt"]))
            next_ugt = st.text_area("Что нужно для следующего УГТ", value=sup_df.iloc[0]["supervisor_next_ugt"])
            if st.button("Сохранить данные руководителя"):
                mask = df["supervisor"] == selected_sup
                if mask.any():
                    df.loc[mask, "supervisor_department"] = department
                    df.loc[mask, "supervisor_grnti"] = grnti
                    df.loc[mask, "supervisor_team"] = team
                    df.loc[mask, "supervisor_competencies"] = comp
                    df.loc[mask, "supervisor_publications"] = pubs
                    df.loc[mask, "supervisor_grants"] = grants
                    df.loc[mask, "supervisor_niokr"] = niokr
                    df.loc[mask, "supervisor_rid_protected"] = rid_protected
                    df.loc[mask, "supervisor_rid_unprotected"] = rid_unprotected
                    df.loc[mask, "supervisor_ugt"] = ugt
                    df.loc[mask, "supervisor_next_ugt"] = next_ugt
                    save_data(df)
                    st.success("Данные сохранены")
                    st.rerun()

        st.subheader("Проекты руководителя")
        if not sup_df.empty:
            proj_table = sup_df[["project_name", "customer", "lifecycle_stage", "sales_stage", "funding_source"]].copy()
            proj_table.columns = ["Название проекта", "Заказчик", "Стадия ЖЦ", "Этап продвижения", "Источник финансирования"]
            st.dataframe(proj_table, hide_index=True, use_container_width=True)

        current_ugt_val = sup_df.iloc[0]["supervisor_ugt"]
        st.subheader(f"УГТ задела научной группы: {current_ugt_val}")
        st.progress(current_ugt_val / 9.0)
        next_info = sup_df.iloc[0]["supervisor_next_ugt"]
        if next_info:
            st.caption(f"📌 Что нужно для следующего УГТ: {next_info}")

        for _, row in sup_df.iterrows():
            with st.expander(f"{row['project_name']} – {row['customer']}"):
                st.write(f"**Решаемая проблема:** {row['problem']}")
                st.write(f"**Конкурент:** {row['competitor']} → **Преимущество:** {row['advantage']}")
                st.write(f"**Партнёр:** {row['partner']} | **Роль МИЭМ:** {row['role_miem']}")
                st.write(f"**Горизонт:** {row['horizon']} | **Источник финансирования:** {row['funding_source']}")
                st.write(f"**Стадия ЖЦ:** {row['lifecycle_stage']} | **Этап продвижения:** {row['sales_stage']}")

# ========== СТРАНИЦА "ПРОЕКТЫ" ==========
elif page == "📋 Проекты":
    st.header("Проекты и заказчики")
    df = load_data()
    if df.empty:
        st.info("Нет проектов. Добавьте первый через форму ниже.")
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            sup_filter = st.selectbox("Научный руководитель", ["Все"] + sorted(df["supervisor"].dropna().unique()))
        with col2:
            sales_filter = st.selectbox("Этап продвижения", ["Все"] + SALES_STAGES)
        with col3:
            lc_filter = st.selectbox("Стадия ЖЦ", ["Все"] + LIFECYCLE_STAGES)

        filtered = df.copy()
        if sup_filter != "Все":
            filtered = filtered[filtered["supervisor"] == sup_filter]
        if sales_filter != "Все":
            filtered = filtered[filtered["sales_stage"] == sales_filter]
        if lc_filter != "Все":
            filtered = filtered[filtered["lifecycle_stage"] == lc_filter]

        st.subheader(f"Всего проектов: {len(filtered)}")
        if not filtered.empty:
            display = filtered[["id", "project_name", "customer", "supervisor", "supervisor_ugt", "lifecycle_stage", "sales_stage", "funding_source"]].copy()
            display.columns = ["ID", "Название проекта", "Заказчик", "Научный руководитель", "УГТ задела", "Стадия ЖЦ", "Этап продвижения", "Источник финансирования"]
            st.dataframe(display, hide_index=True, use_container_width=True)

        # Таблица последних изменений (перенесена с дашборда)
        st.subheader("📜 Последние изменения проектов")
        if "stage_change_date" in df.columns and not df["stage_change_date"].isna().all():
            history = df[["project_name", "customer", "sales_stage", "stage_change_date", "stage_change_reason"]].dropna(subset=["stage_change_date"])
            history = history.sort_values("stage_change_date", ascending=False).head(10)
            history.columns = ["Проект", "Заказчик", "Этап", "Дата", "Причина"]
            st.dataframe(history, hide_index=True, use_container_width=True)
        else:
            st.info("Нет записей об изменениях")

    # Добавление нового проекта
    with st.expander("➕ Добавить новый проект"):
        with st.form("new_project"):
            sup = st.text_input("ФИО научного руководителя*")
            proj = st.text_input("Название проекта*")
            cust = st.text_input("Заказчик*")
            prob = st.text_area("Решаемая проблема")
            comp = st.text_input("Конкурент")
            adv = st.text_input("Преимущество")
            part = st.text_input("Партнёр")
            role = st.selectbox("Роль МИЭМ", [""] + ROLES_MIEM)
            hor = st.selectbox("Горизонт реализации", [""] + HORIZON_LIST)
            funding = st.selectbox("Источник финансирования", FUNDING_SOURCES)
            lc = st.selectbox("Стадия ЖЦ", LIFECYCLE_STAGES)
            sales = st.selectbox("Этап продвижения", SALES_STAGES)
            submitted = st.form_submit_button("Создать")
            if submitted:
                if sup and proj and cust:
                    new_id = get_next_id(df)
                    new_row = pd.DataFrame([{
                        "id": new_id, "supervisor": sup,
                        "supervisor_department": "", "supervisor_grnti": "", "supervisor_team": "",
                        "supervisor_competencies": "", "supervisor_publications": "", "supervisor_grants": "",
                        "supervisor_niokr": "", "supervisor_rid_protected": "", "supervisor_rid_unprotected": "",
                        "supervisor_ugt": 1, "supervisor_next_ugt": "",
                        "project_name": proj, "customer": cust, "problem": prob,
                        "competitor": comp, "advantage": adv, "partner": part, "role_miem": role,
                        "horizon": hor, "funding_source": funding,
                        "lifecycle_stage": lc, "sales_stage": sales,
                        "stage_change_reason": "Создание проекта", "stage_change_date": datetime.now()
                    }])
                    df = pd.concat([df, new_row], ignore_index=True)
                    save_data(df)
                    st.success("Проект добавлен")
                    st.rerun()
                else:
                    st.error("Заполните все поля со звёздочкой")

    # Редактирование существующего проекта
    if not df.empty:
        with st.expander("✏️ Редактировать проект"):
            # Создаём словарь для выбора по ID + Название
            project_dict = {f"{row['id']} – {row['project_name']}": row['id'] for _, row in df.iterrows()}
            selected_label = st.selectbox("Выберите проект", list(project_dict.keys()))
            sel_id = project_dict[selected_label]
            proj = df[df["id"] == sel_id].iloc[0]
            with st.form("edit_project"):
                sup = st.text_input("ФИО руководителя", value=proj["supervisor"])
                pname = st.text_input("Название", value=proj["project_name"])
                cust = st.text_input("Заказчик", value=proj["customer"])
                prob = st.text_area("Проблема", value=proj["problem"])
                comp = st.text_input("Конкурент", value=proj["competitor"])
                adv = st.text_input("Преимущество", value=proj["advantage"])
                part = st.text_input("Партнёр", value=proj["partner"])
                role = st.selectbox("Роль МИЭМ", ROLES_MIEM, index=ROLES_MIEM.index(proj["role_miem"]) if proj["role_miem"] in ROLES_MIEM else 0)
                hor = st.selectbox("Горизонт", HORIZON_LIST, index=HORIZON_LIST.index(proj["horizon"]) if proj["horizon"] in HORIZON_LIST else 0)
                funding = st.selectbox("Источник", FUNDING_SOURCES, index=FUNDING_SOURCES.index(proj["funding_source"]) if proj["funding_source"] in FUNDING_SOURCES else 0)
                lc = st.selectbox("Стадия ЖЦ", LIFECYCLE_STAGES, index=LIFECYCLE_STAGES.index(proj["lifecycle_stage"]) if proj["lifecycle_stage"] in LIFECYCLE_STAGES else 0)
                sales = st.selectbox("Этап продаж", SALES_STAGES, index=SALES_STAGES.index(proj["sales_stage"]) if proj["sales_stage"] in SALES_STAGES else 0)
                col1, col2 = st.columns(2)
                saved = col1.form_submit_button("Сохранить")
                deleted = col2.form_submit_button("Удалить")
                if saved:
                    df.loc[df["id"] == sel_id, ["supervisor", "project_name", "customer", "problem", "competitor",
                                                 "advantage", "partner", "role_miem", "horizon", "funding_source",
                                                 "lifecycle_stage", "sales_stage"]] = \
                        [sup, pname, cust, prob, comp, adv, part, role, hor, funding, lc, sales]
                    df.loc[df["id"] == sel_id, "stage_change_date"] = datetime.now()
                    df.loc[df["id"] == sel_id, "stage_change_reason"] = "Редактирование"
                    save_data(df)
                    st.success("Сохранено")
                    st.rerun()
                if deleted:
                    df = df[df["id"] != sel_id]
                    save_data(df)
                    st.success("Удалено")
                    st.rerun()

# ========== СТРАНИЦА "ДАШБОРД" ==========
elif page == "📊 Дашборд":
    st.header("Аналитика и воронка продвижения")
    df = load_data()
    if df.empty:
        st.info("Нет данных для анализа")
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric("Всего проектов", len(df))
        col2.metric("Научных руководителей", df["supervisor"].nunique())
        col3.metric("Средний УГТ задела", f"{df['supervisor_ugt'].mean():.1f}")

        st.subheader("Воронка продвижения")
        sales_counts = df["sales_stage"].value_counts().reindex(SALES_STAGES, fill_value=0).reset_index()
        sales_counts.columns = ["Этап продвижения", "Количество проектов"]
        mean_sales = sales_counts["Количество проектов"].mean()
        base = alt.Chart(sales_counts).mark_bar(color="steelblue").encode(
            x=alt.X("Этап продвижения", sort=SALES_STAGES, title="Этап продвижения", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            y=alt.Y("Количество проектов", title="Количество", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold'))
        ).properties(width=700, height=400)
        line = alt.Chart(pd.DataFrame({'y': [mean_sales]})).mark_rule(strokeDash=[5,5], color='gray').encode(y='y')
        text = alt.Chart(pd.DataFrame({'y': [mean_sales], 'txt': [f'среднее = {mean_sales:.1f}']})).mark_text(dy=-10, color='gray', fontSize=12).encode(y='y', text='txt')
        st.altair_chart(base + line + text, use_container_width=True)

        st.subheader("Стадии жизненного цикла продукта")
        lc_counts = df["lifecycle_stage"].value_counts().reindex(LIFECYCLE_STAGES, fill_value=0).reset_index()
        lc_counts.columns = ["Стадия ЖЦ", "Количество проектов"]
        mean_lc = lc_counts["Количество проектов"].mean()
        base2 = alt.Chart(lc_counts).mark_bar(color="darkorange").encode(
            x=alt.X("Стадия ЖЦ", sort=LIFECYCLE_STAGES, title="Стадия ЖЦ", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            y=alt.Y("Количество проектов", title="Количество", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold'))
        ).properties(width=700, height=400)
        line2 = alt.Chart(pd.DataFrame({'y': [mean_lc]})).mark_rule(strokeDash=[5,5], color='gray').encode(y='y')
        text2 = alt.Chart(pd.DataFrame({'y': [mean_lc], 'txt': [f'среднее = {mean_lc:.1f}']})).mark_text(dy=-10, color='gray', fontSize=12).encode(y='y', text='txt')
        st.altair_chart(base2 + line2 + text2, use_container_width=True)

        st.subheader("Уровень УГТ задела научных групп")
        ugt_counts = df["supervisor_ugt"].value_counts().sort_index().reset_index()
        ugt_counts.columns = ["УГТ", "Количество руководителей"]
        ugt_counts["УГТ"] = ugt_counts["УГТ"].astype(int)
        mean_ugt = ugt_counts["Количество руководителей"].mean()
        base3 = alt.Chart(ugt_counts).mark_bar(color="green").encode(
            x=alt.X("УГТ:Q", title="УГТ", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold', format='d')),
            y=alt.Y("Количество руководителей:Q", title="Количество", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold'))
        ).properties(width=600, height=400)
        line3 = alt.Chart(pd.DataFrame({'y': [mean_ugt]})).mark_rule(strokeDash=[5,5], color='gray').encode(y='y')
        text3 = alt.Chart(pd.DataFrame({'y': [mean_ugt], 'txt': [f'среднее = {mean_ugt:.1f}']})).mark_text(dy=-10, color='gray', fontSize=12).encode(y='y', text='txt')
        st.altair_chart(base3 + line3 + text3, use_container_width=True)

# ========== СТРАНИЦА "ЭКСПОРТ / ИМПОРТ" ==========
elif page == "💾 Экспорт / Импорт":
    st.header("Резервное копирование и восстановление")
    df = load_data()

    st.subheader("📥 Сохранить текущую базу")
    if not df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button(
            label="Сохранить projects.xlsx",
            data=output.getvalue(),
            file_name="projects_backup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Нет данных для экспорта")

    st.divider()
    st.subheader("📤 Восстановить из файла")
    uploaded = st.file_uploader("Загрузите ранее сохранённый .xlsx", type=["xlsx"])
    if uploaded is not None:
        try:
            backup_df = pd.read_excel(uploaded)
            required = ["id", "supervisor", "project_name", "customer", "supervisor_ugt"]
            if all(c in backup_df.columns for c in required):
                backup_df = backup_df.reset_index(drop=True)
                backup_df.to_excel(DATA_FILE, index=False)
                st.success("База восстановлена! Перезагрузите страницу.")
                st.rerun()
            else:
                st.error(f"Файл не содержит колонок: {required}")
        except Exception as e:
            st.error(f"Ошибка: {e}")
