import streamlit as st
import pandas as pd
import os
from datetime import datetime
import re
import altair as alt

st.set_page_config(page_title="CRM МИЭМ НАУКА", layout="wide")
st.title("🏛️ CRM МИЭМ НАУКА — Управление научными проектами")

DATA_FILE = "projects.xlsx"

# ---------- НОВЫЙ СПИСОК КОЛОНОК ----------
COLUMNS = [
    "id", 
    "supervisor", 
    "supervisor_competencies", 
    "supervisor_publications", 
    "supervisor_past_projects",       # вместо грантов - выполненные проекты
    "supervisor_team",                # команда (количество, состав, роли)
    "supervisor_grnti",               # шифр ГРНТИ
    "supervisor_ugt",                 # УГТ руководителя (1-9)
    "supervisor_barriers",            # барьеры для повышения УГТ
    "project_name",
    "customer",
    "problem",                        # решаемая проблема заказчика
    "competitor",
    "advantage",
    "partner",
    "role_miem",
    "horizon",
    "funding_source",                 # источник финансирования проекта
    "lifecycle_stage",
    "sales_stage",                    # этап продвижения
    "department",
    "stage_change_reason",
    "stage_change_date"
]

def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, dtype={"id": int})
    else:
        df = pd.DataFrame(columns=COLUMNS)
        df["id"] = df["id"].astype(int)

    # Добавляем недостающие колонки (для миграции)
    for col in COLUMNS:
        if col not in df.columns:
            if col == "id":
                df[col] = 0
            elif col == "supervisor_ugt":
                df[col] = 1
            else:
                df[col] = ""

    # Приводим текстовые колонки к строковому типу
    text_cols = [col for col in COLUMNS if col not in ["id", "supervisor_ugt", "stage_change_date"]]
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

# ---------- СПИСКИ ДЛЯ ВЫБОРА (строгий порядок) ----------
LIFECYCLE_STAGES = ["Планирование (НИР)", "Проектирование (ОКР)", "Разработка", "Внедрение", "Эксплуатация"]
SALES_STAGES = [
    "Квалификация", "Выявление проблем", "Формирование видения",
    "Обоснование ценности", "Проработка решения", "Презентация",
    "Переговоры и возражения", "Закрытие сделки", "Поддержка и развитие"
]
ROLES_MIEM = ["Субподрядчик", "Соисполнитель", "Лицензиар", "Сервисный центр", "Технологический аудитор",
              "Консультант", "Другая роль"]
BARRIERS_LIST = ["Нет оформленных прав на РИД", "Нет прототипа в реальных условиях", "Нет подходящего партнёра",
                 "Нет времени", "Нет исполнителей", "Нет понимания рынка", "Нет коммерческого потенциала",
                 "Отсутствие инфраструктуры", "Другой барьер"]
HORIZON_LIST = ["0-3 месяца", "3-6 месяцев", "6-12 месяцев", "1-3 года", "Другой срок"]
FUNDING_SOURCES = ["Внутренний (грант ВШЭ/МИЭМ)", "Внешний институциональный (РНФ, РФФИ, Минобр)",
                   "Внешний корпоративный (компания)", "Смешанный", "Другое"]

# ---------- НАВИГАЦИЯ ----------
st.sidebar.title("Навигация")
page = st.sidebar.radio("Перейти", ["📋 Проекты", "👨‍🔬 Научные руководители", "📊 Дашборд"])

# ---------- СТРАНИЦА "ПРОЕКТЫ" ----------
if page == "📋 Проекты":
    st.header("Проекты и заказчики")
    df = load_data()

    # Фильтры
    col1, col2, col3 = st.columns(3)
    with col1:
        supervisor_filter = st.selectbox("Научный руководитель", ["Все"] + sorted(df["supervisor"].dropna().unique().tolist()) if not df.empty else ["Все"])
    with col2:
        sales_filter = st.selectbox("Этап продвижения", ["Все"] + SALES_STAGES)
    with col3:
        lc_filter = st.selectbox("Стадия ЖЦ", ["Все"] + LIFECYCLE_STAGES)

    filtered_df = df.copy()
    if supervisor_filter != "Все" and not df.empty:
        filtered_df = filtered_df[filtered_df["supervisor"] == supervisor_filter]
    if sales_filter != "Все" and not df.empty:
        filtered_df = filtered_df[filtered_df["sales_stage"] == sales_filter]
    if lc_filter != "Все" and not df.empty:
        filtered_df = filtered_df[filtered_df["lifecycle_stage"] == lc_filter]

    st.subheader(f"Всего проектов: {len(filtered_df)}")
    if not filtered_df.empty:
        st.dataframe(
            filtered_df[["id", "project_name", "customer", "supervisor", "lifecycle_stage", "sales_stage", "funding_source"]],
            hide_index=True,
            use_container_width=True
        )
    else:
        st.info("Нет проектов по выбранным фильтрам")

    # Добавление нового проекта
    with st.expander("➕ Добавить новый проект"):
        with st.form("new_project"):
            sup = st.text_input("ФИО научного руководителя*")
            proj = st.text_input("Название проекта*")
            cust = st.text_input("Заказчик*")
            prob = st.text_area("Решаемая проблема заказчика")
            comp = st.text_input("Конкурент")
            adv = st.text_input("Преимущество над конкурентом")
            part = st.text_input("Партнёр")
            role = st.selectbox("Роль МИЭМ в цепочке", [""] + ROLES_MIEM)
            hor = st.selectbox("Горизонт реализации", [""] + HORIZON_LIST)
            funding = st.selectbox("Источник финансирования проекта", FUNDING_SOURCES)
            lc_stage = st.selectbox("Стадия ЖЦ продукта", LIFECYCLE_STAGES)
            sales_stage = st.selectbox("Этап продвижения", SALES_STAGES)
            dept = st.text_input("Подразделение МИЭМ")
            submitted = st.form_submit_button("Создать")
            if submitted:
                if sup and proj and cust:
                    new_id = get_next_id(df)
                    new_row = pd.DataFrame([{
                        "id": new_id,
                        "supervisor": sup,
                        "supervisor_competencies": "", "supervisor_publications": "", "supervisor_past_projects": "",
                        "supervisor_team": "", "supervisor_grnti": "", "supervisor_ugt": 1, "supervisor_barriers": "",
                        "project_name": proj, "customer": cust, "problem": prob,
                        "competitor": comp, "advantage": adv, "partner": part, "role_miem": role,
                        "horizon": hor, "funding_source": funding,
                        "lifecycle_stage": lc_stage, "sales_stage": sales_stage, "department": dept,
                        "stage_change_reason": "Создание проекта", "stage_change_date": datetime.now()
                    }])
                    df = pd.concat([df, new_row], ignore_index=True)
                    save_data(df)
                    st.success("Проект добавлен")
                    st.rerun()
                else:
                    st.error("Заполните ФИО руководителя, название проекта и заказчика")

    # Редактирование существующего проекта
    if not df.empty:
        with st.expander("✏️ Редактировать существующий проект"):
            selected_id = st.selectbox("Выберите ID проекта", df["id"].tolist())
            project = df[df["id"] == selected_id].iloc[0]
            with st.form("edit_project"):
                sup = st.text_input("ФИО научного руководителя", value=project["supervisor"])
                proj = st.text_input("Название проекта", value=project["project_name"])
                cust = st.text_input("Заказчик", value=project["customer"])
                prob = st.text_area("Решаемая проблема", value=project["problem"])
                comp = st.text_input("Конкурент", value=project["competitor"])
                adv = st.text_input("Преимущество", value=project["advantage"])
                part = st.text_input("Партнёр", value=project["partner"])
                role = st.selectbox("Роль МИЭМ", ROLES_MIEM, index=ROLES_MIEM.index(project["role_miem"]) if project["role_miem"] in ROLES_MIEM else 0)
                hor = st.selectbox("Горизонт реализации", HORIZON_LIST, index=HORIZON_LIST.index(project["horizon"]) if project["horizon"] in HORIZON_LIST else 0)
                funding = st.selectbox("Источник финансирования", FUNDING_SOURCES, index=FUNDING_SOURCES.index(project["funding_source"]) if project["funding_source"] in FUNDING_SOURCES else 0)
                lc_stage = st.selectbox("Стадия ЖЦ", LIFECYCLE_STAGES, index=LIFECYCLE_STAGES.index(project["lifecycle_stage"]) if project["lifecycle_stage"] in LIFECYCLE_STAGES else 0)
                sales_stage = st.selectbox("Этап продвижения", SALES_STAGES, index=SALES_STAGES.index(project["sales_stage"]) if project["sales_stage"] in SALES_STAGES else 0)
                dept = st.text_input("Подразделение", value=project["department"])
                col1, col2 = st.columns(2)
                with col1:
                    saved = st.form_submit_button("Сохранить")
                with col2:
                    deleted = st.form_submit_button("Удалить проект")
                if saved:
                    df.loc[df["id"] == selected_id, ["supervisor", "project_name", "customer", "problem", "competitor",
                                                     "advantage", "partner", "role_miem", "horizon", "funding_source",
                                                     "lifecycle_stage", "sales_stage", "department"]] = \
                        [sup, proj, cust, prob, comp, adv, part, role, hor, funding, lc_stage, sales_stage, dept]
                    df.loc[df["id"] == selected_id, "stage_change_date"] = datetime.now()
                    df.loc[df["id"] == selected_id, "stage_change_reason"] = "Редактирование проекта"
                    save_data(df)
                    st.success("Сохранено")
                    st.rerun()
                if deleted:
                    df = df[df["id"] != selected_id]
                    save_data(df)
                    st.success("Удалено")
                    st.rerun()

# ---------- СТРАНИЦА "НАУЧНЫЕ РУКОВОДИТЕЛИ" ----------
elif page == "👨‍🔬 Научные руководители":
    st.header("Научные руководители и их проекты")
    df = load_data()
    if df.empty:
        st.info("Нет данных. Добавьте проекты на вкладке 'Проекты'.")
    else:
        supervisors = df["supervisor"].dropna().unique()
        selected_sup = st.selectbox("Выберите научного руководителя", supervisors)
        sup_df = df[df["supervisor"] == selected_sup]
        st.subheader(f"Проекты руководителя {selected_sup}")

        with st.expander("📋 Карточка научного руководителя (задел)"):
            # Команда
            team = st.text_area("Команда (количество, состав, роли)", value=sup_df.iloc[0]["supervisor_team"] if "supervisor_team" in sup_df.columns else "")
            # Шифр ГРНТИ
            grnti = st.text_input("Шифр ГРНТИ", value=sup_df.iloc[0]["supervisor_grnti"] if "supervisor_grnti" in sup_df.columns else "")
            # УГТ руководителя
            sup_ugt = st.slider("Уровень готовности технологии (УГТ) руководителя", 1, 9, value=int(sup_df.iloc[0]["supervisor_ugt"]) if "supervisor_ugt" in sup_df.columns else 1)
            # Барьеры повышения УГТ
            barriers = st.text_area("Барьеры для повышения УГТ", value=sup_df.iloc[0]["supervisor_barriers"] if "supervisor_barriers" in sup_df.columns else "")
            # Ключевые компетенции
            comp = st.text_area("Ключевые компетенции", value=sup_df.iloc[0]["supervisor_competencies"] if "supervisor_competencies" in sup_df.columns else "")
            # Публикации (до 5)
            pubs = st.text_area("Публикации (до 5)", value=sup_df.iloc[0]["supervisor_publications"] if "supervisor_publications" in sup_df.columns else "")
            # Ранее выполненные проекты (НИОКР, гранты, внедрения)
            past = st.text_area("Ранее выполненные проекты (гранты, НИОКР, внедрения)", value=sup_df.iloc[0]["supervisor_past_projects"] if "supervisor_past_projects" in sup_df.columns else "")
            if st.button("Сохранить данные руководителя"):
                mask = df["supervisor"] == selected_sup
                if mask.any():
                    df.loc[mask, "supervisor_team"] = team
                    df.loc[mask, "supervisor_grnti"] = grnti
                    df.loc[mask, "supervisor_ugt"] = sup_ugt
                    df.loc[mask, "supervisor_barriers"] = barriers
                    df.loc[mask, "supervisor_competencies"] = comp
                    df.loc[mask, "supervisor_publications"] = pubs
                    df.loc[mask, "supervisor_past_projects"] = past
                    save_data(df)
                    st.success("Сохранено")
                    st.rerun()
                else:
                    st.error("Руководитель не найден")

        # Таблица проектов руководителя
        st.subheader("Проекты руководителя")
        st.dataframe(
            sup_df[["project_name", "customer", "lifecycle_stage", "sales_stage", "funding_source"]],
            hide_index=True,
            use_container_width=True
        )

        # Прогресс по УГТ (средний по проектам не нужен, используем УГТ руководителя)
        current_ugt = sup_df.iloc[0]["supervisor_ugt"] if "supervisor_ugt" in sup_df.columns else 1
        st.subheader(f"УГТ руководителя: {current_ugt}")
        st.progress(current_ugt / 9.0)

        # Детализация проектов
        for _, row in sup_df.iterrows():
            with st.expander(f"{row['project_name']} – {row['customer']}"):
                st.write(f"**Решаемая проблема:** {row['problem']}")
                st.write(f"**Конкурент:** {row['competitor']} → **Преимущество:** {row['advantage']}")
                st.write(f"**Партнёр:** {row['partner']} | **Роль МИЭМ:** {row['role_miem']}")
                st.write(f"**Горизонт реализации:** {row['horizon']} | **Источник финансирования:** {row['funding_source']}")
                st.write(f"**Стадия ЖЦ продукта:** {row['lifecycle_stage']} | **Этап продвижения:** {row['sales_stage']}")

# ---------- СТРАНИЦА "ДАШБОРД" ----------
elif page == "📊 Дашборд":
    st.header("Аналитика и воронка продвижения")
    df = load_data()
    if df.empty:
        st.info("Нет данных для анализа")
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric("Всего проектов", len(df))
        col2.metric("Научных руководителей", df["supervisor"].nunique())
        col3.metric("Средний УГТ руководителей", f"{df['supervisor_ugt'].mean():.1f}")

        # Воронка продаж (этапы продвижения) с яркими цветами через Altair
        st.subheader("Воронка этапов продвижения")
        sales_counts = df["sales_stage"].value_counts().reindex(SALES_STAGES, fill_value=0).reset_index()
        sales_counts.columns = ["Этап продвижения", "Количество проектов"]
        chart = alt.Chart(sales_counts).mark_bar(color="steelblue").encode(
            x=alt.X("Этап продвижения", sort=SALES_STAGES, title="Этап продвижения"),
            y=alt.Y("Количество проектов", title="Количество проектов")
        ).properties(width=700, height=400)
        st.altair_chart(chart, use_container_width=True)

        # Распределение по стадиям ЖЦ
        st.subheader("Стадии жизненного цикла продукта")
        lc_counts = df["lifecycle_stage"].value_counts().reindex(LIFECYCLE_STAGES, fill_value=0).reset_index()
        lc_counts.columns = ["Стадия ЖЦ", "Количество проектов"]
        chart2 = alt.Chart(lc_counts).mark_bar(color="darkorange").encode(
            x=alt.X("Стадия ЖЦ", sort=LIFECYCLE_STAGES, title="Стадия ЖЦ"),
            y=alt.Y("Количество проектов", title="Количество проектов")
        ).properties(width=700, height=400)
        st.altair_chart(chart2, use_container_width=True)

        # Распределение по УГТ руководителей
        st.subheader("Уровень УГТ научных руководителей")
        ugt_counts = df["supervisor_ugt"].value_counts().sort_index().reset_index()
        ugt_counts.columns = ["УГТ", "Количество руководителей"]
        chart3 = alt.Chart(ugt_counts).mark_bar(color="green").encode(
            x=alt.X("УГТ:Q", title="УГТ"),
            y=alt.Y("Количество руководителей:Q", title="Количество")
        )
        st.altair_chart(chart3, use_container_width=True)

        # Последние изменения
        st.subheader("Последние изменения проектов")
        if "stage_change_date" in df.columns and not df["stage_change_date"].isna().all():
            history = df[["project_name", "customer", "sales_stage", "stage_change_date", "stage_change_reason"]].dropna(subset=["stage_change_date"])
            history = history.sort_values("stage_change_date", ascending=False).head(10)
            st.dataframe(history, hide_index=True, use_container_width=True)
