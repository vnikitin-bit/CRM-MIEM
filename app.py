import streamlit as st
import pandas as pd
import os
from datetime import datetime
import re

st.set_page_config(page_title="CRM МИЭМ НАУКА", layout="wide")
st.title("🏛️ CRM НИУ ВШЭ МИЭМ НАУКА — Управление научными проектами")

DATA_FILE = "projects.xlsx"

# Определяем все колонки
COLUMNS = [
    "id", "supervisor", "supervisor_competencies", "supervisor_publications", "supervisor_grants",
    "project_name", "customer", "problem", "competitor", "advantage", "partner", "role_miem",
    "barriers", "horizon", "ugt", "lifecycle_stage", "sales_stage", "department",
    "stage_change_reason", "stage_change_date"
]

def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE)
        # Убеждаемся, что все колонки есть
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = "" if col not in ["id", "ugt"] else (0 if col == "id" else 1)
    else:
        df = pd.DataFrame(columns=COLUMNS)
        df["id"] = df["id"].astype(int)
    # Приведение типов
    if "id" in df.columns:
        df["id"] = pd.to_numeric(df["id"], errors="coerce").fillna(0).astype(int)
    if "ugt" in df.columns:
        df["ugt"] = pd.to_numeric(df["ugt"], errors="coerce").fillna(1).astype(int)
    if "stage_change_date" in df.columns:
        df["stage_change_date"] = pd.to_datetime(df["stage_change_date"], errors="coerce")
    # Удаляем записи с пустым заказчиком или "[вручную]"
    if "customer" in df.columns:
        df = df[~df["customer"].astype(str).str.contains(r"\[вручную\]", na=False, case=False)]
        df = df[df["customer"].notna() & (df["customer"].astype(str).str.strip() != "")]
    df = df.reset_index(drop=True)
    if not df.empty:
        df = df.sort_values("id").reset_index(drop=True)
    return df

def save_data(df):
    # Принудительно сохраняем только нужные колонки
    df_to_save = df[COLUMNS].copy()
    df_to_save.to_excel(DATA_FILE, index=False)

def get_next_id(df):
    return int(df["id"].max() + 1) if not df.empty else 1

# Списки для выбора
LIFECYCLE_STAGES = ["Планирование (НИР)", "Проектирование (ОКР)", "Разработка", "Внедрение", "Эксплуатация"]
SALES_STAGES = ["Квалификация", "Выявление проблем", "Формирование видения", "Обоснование ценности",
                "Проработка решения", "Презентация", "Переговоры и возражения", "Закрытие сделки",
                "Поддержка и развитие"]
ROLES_MIEM = ["Субподрядчик", "Соисполнитель", "Лицензиар", "Сервисный центр", "Технологический аудитор",
              "Консультант", "Другая роль"]
BARRIERS_LIST = ["Нет оформленных прав на РИД", "Нет прототипа в реальных условиях", "Нет подходящего партнёра",
                 "Нет времени", "Нет исполнителей", "Нет понимания рынка", "Нет коммерческого потенциала",
                 "Отсутствие инфраструктуры", "Другой барьер"]
HORIZON_LIST = ["0-3 месяца", "3-6 месяцев", "6-12 месяцев", "1-3 года", "Другой срок"]

# Боковая навигация
st.sidebar.title("Навигация")
page = st.sidebar.radio("Перейти", ["📋 Проекты", "👨‍🔬 Научные руководители", "📊 Дашборд"])

# Страница "Проекты"
if page == "📋 Проекты":
    st.header("Проекты и заказчики")
    df = load_data()

    # Фильтры (безопасно, если df пуст)
    if df.empty:
        supervisors_list = []
        ugt_list = []
    else:
        supervisors_list = sorted(df["supervisor"].dropna().unique().tolist())
        ugt_list = sorted(df["ugt"].unique())
    col1, col2, col3 = st.columns(3)
    with col1:
        supervisor_filter = st.selectbox("Научный руководитель", ["Все"] + supervisors_list)
    with col2:
        ugt_filter = st.selectbox("УГТ", ["Все"] + ugt_list)
    with col3:
        sales_filter = st.selectbox("Этап продаж", ["Все"] + SALES_STAGES)

    filtered_df = df.copy() if not df.empty else df
    if supervisor_filter != "Все" and not df.empty:
        filtered_df = filtered_df[filtered_df["supervisor"] == supervisor_filter]
    if ugt_filter != "Все" and not df.empty:
        filtered_df = filtered_df[filtered_df["ugt"] == ugt_filter]
    if sales_filter != "Все" and not df.empty:
        filtered_df = filtered_df[filtered_df["sales_stage"] == sales_filter]

    st.subheader(f"Всего проектов: {len(filtered_df)}")
    if not filtered_df.empty:
        st.dataframe(
            filtered_df[["id", "project_name", "customer", "supervisor", "ugt", "lifecycle_stage", "sales_stage"]],
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
            prob = st.text_area("Решаемая проблема")
            comp = st.text_input("Конкурент")
            adv = st.text_input("Преимущество над конкурентом")
            part = st.text_input("Партнёр")
            role = st.selectbox("Роль МИЭМ в цепочке", [""] + ROLES_MIEM)
            barr = st.selectbox("Барьеры", [""] + BARRIERS_LIST)
            hor = st.selectbox("Горизонт реализации", [""] + HORIZON_LIST)
            ugt = st.slider("УГТ", 1, 9, 4)
            lc_stage = st.selectbox("Стадия ЖЦ продукта", LIFECYCLE_STAGES)
            sales_stage = st.selectbox("Этап продаж (Solution Selling)", SALES_STAGES)
            dept = st.text_input("Подразделение МИЭМ")
            submitted = st.form_submit_button("Создать")
            if submitted:
                if sup and proj and cust:
                    new_id = get_next_id(df)
                    new_row = pd.DataFrame([{
                        "id": new_id,
                        "supervisor": sup,
                        "supervisor_competencies": "",
                        "supervisor_publications": "",
                        "supervisor_grants": "",
                        "project_name": proj,
                        "customer": cust,
                        "problem": prob,
                        "competitor": comp,
                        "advantage": adv,
                        "partner": part,
                        "role_miem": role,
                        "barriers": barr,
                        "horizon": hor,
                        "ugt": ugt,
                        "lifecycle_stage": lc_stage,
                        "sales_stage": sales_stage,
                        "department": dept,
                        "stage_change_reason": "Создание проекта",
                        "stage_change_date": datetime.now()
                    }])
                    if df.empty:
                        df = new_row
                    else:
                        df = pd.concat([df, new_row], ignore_index=True)
                    save_data(df)
                    st.success("Проект добавлен!")
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
                barr = st.selectbox("Барьеры", BARRIERS_LIST, index=BARRIERS_LIST.index(project["barriers"]) if project["barriers"] in BARRIERS_LIST else 0)
                hor = st.selectbox("Горизонт реализации", HORIZON_LIST, index=HORIZON_LIST.index(project["horizon"]) if project["horizon"] in HORIZON_LIST else 0)
                ugt = st.slider("УГТ", 1, 9, int(project["ugt"]))
                lc_stage = st.selectbox("Стадия ЖЦ", LIFECYCLE_STAGES, index=LIFECYCLE_STAGES.index(project["lifecycle_stage"]) if project["lifecycle_stage"] in LIFECYCLE_STAGES else 0)
                sales_stage = st.selectbox("Этап продаж", SALES_STAGES, index=SALES_STAGES.index(project["sales_stage"]) if project["sales_stage"] in SALES_STAGES else 0)
                dept = st.text_input("Подразделение", value=project["department"])
                col1, col2 = st.columns(2)
                with col1:
                    saved = st.form_submit_button("Сохранить")
                with col2:
                    deleted = st.form_submit_button("Удалить проект")
                if saved:
                    # Обновляем запись
                    idx = df[df["id"] == selected_id].index[0]
                    df.loc[idx, "supervisor"] = sup
                    df.loc[idx, "project_name"] = proj
                    df.loc[idx, "customer"] = cust
                    df.loc[idx, "problem"] = prob
                    df.loc[idx, "competitor"] = comp
                    df.loc[idx, "advantage"] = adv
                    df.loc[idx, "partner"] = part
                    df.loc[idx, "role_miem"] = role
                    df.loc[idx, "barriers"] = barr
                    df.loc[idx, "horizon"] = hor
                    df.loc[idx, "ugt"] = ugt
                    df.loc[idx, "lifecycle_stage"] = lc_stage
                    df.loc[idx, "sales_stage"] = sales_stage
                    df.loc[idx, "department"] = dept
                    df.loc[idx, "stage_change_date"] = datetime.now()
                    df.loc[idx, "stage_change_reason"] = "Редактирование проекта"
                    save_data(df)
                    st.success("Сохранено")
                    st.rerun()
                if deleted:
                    df = df[df["id"] != selected_id]
                    save_data(df)
                    st.success("Удалено")
                    st.rerun()

# Страница "Научные руководители"
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

        with st.expander("Карточка руководителя (компетенции, гранты, публикации)"):
            comp = st.text_area("Ключевые компетенции", value=sup_df.iloc[0]["supervisor_competencies"] if "supervisor_competencies" in sup_df.columns else "")
            pubs = st.text_area("Публикации", value=sup_df.iloc[0]["supervisor_publications"] if "supervisor_publications" in sup_df.columns else "")
            grants = st.text_area("Гранты", value=sup_df.iloc[0]["supervisor_grants"] if "supervisor_grants" in sup_df.columns else "")
            if st.button("Сохранить информацию о руководителе"):
                # Обновляем все строки этого руководителя
                idxs = sup_df.index
                df.loc[idxs, "supervisor_competencies"] = comp
                df.loc[idxs, "supervisor_publications"] = pubs
                df.loc[idxs, "supervisor_grants"] = grants
                save_data(df)
                st.success("Сохранено")
                st.rerun()

        st.dataframe(
            sup_df[["project_name", "customer", "ugt", "lifecycle_stage", "sales_stage"]],
            hide_index=True,
            use_container_width=True
        )

        avg_ugt = sup_df["ugt"].mean()
        st.subheader(f"Средний уровень УГТ по проектам: {avg_ugt:.1f}")
        st.progress(avg_ugt / 9.0)

        for _, row in sup_df.iterrows():
            with st.expander(f"{row['project_name']} – {row['customer']} (УГТ {row['ugt']})"):
                st.write(f"**Проблема:** {row['problem']}")
                st.write(f"**Конкурент:** {row['competitor']} → **Преимущество:** {row['advantage']}")
                st.write(f"**Партнёр:** {row['partner']} | **Роль МИЭМ:** {row['role_miem']}")
                st.write(f"**Барьеры:** {row['barriers']} | **Горизонт:** {row['horizon']}")
                st.write(f"**Стадия ЖЦ:** {row['lifecycle_stage']} | **Этап продаж:** {row['sales_stage']}")

# Страница "Дашборд"
elif page == "📊 Дашборд":
    st.header("Аналитика и воронка продаж")
    df = load_data()
    if df.empty:
        st.info("Нет данных для анализа")
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric("Всего проектов", len(df))
        col2.metric("Средний УГТ", f"{df['ugt'].mean():.1f}")
        col3.metric("Научных руководителей", df["supervisor"].nunique())

        st.subheader("Воронка продаж (этапы Solution Selling)")
        sales_counts = df["sales_stage"].value_counts().reindex(SALES_STAGES, fill_value=0)
        st.bar_chart(sales_counts)

        st.subheader("Распределение проектов по УГТ")
        ugt_counts = df["ugt"].value_counts().sort_index()
        st.bar_chart(ugt_counts)

        st.subheader("Стадии жизненного цикла продукта")
        lc_counts = df["lifecycle_stage"].value_counts().reindex(LIFECYCLE_STAGES, fill_value=0)
        st.bar_chart(lc_counts)

        st.subheader("Топ-5 научных руководителей по среднему УГТ")
        sup_ugt = df.groupby("supervisor")["ugt"].mean().sort_values(ascending=False).head(5)
        if not sup_ugt.empty:
            st.dataframe(sup_ugt.reset_index(), hide_index=True, use_container_width=True)

        st.subheader("Последние изменения проектов")
        if "stage_change_date" in df.columns and not df["stage_change_date"].isna().all():
            history = df[["project_name", "customer", "sales_stage", "stage_change_date", "stage_change_reason"]].dropna(subset=["stage_change_date"])
            history = history.sort_values("stage_change_date", ascending=False).head(10)
            st.dataframe(history, hide_index=True, use_container_width=True)
