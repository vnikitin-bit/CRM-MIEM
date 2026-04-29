import streamlit as st
import pandas as pd
import os
from datetime import datetime
import re
import altair as alt
import io

st.set_page_config(page_title="CRM МИЭМ НАУКА", layout="wide")
st.title("🏛️ CRM МИЭМ НАУКА — Управление научными проектами")

DATA_FILE = "projects.xlsx"

# ---------- КОЛОНКИ (обновлены) ----------
COLUMNS = [
    "id",
    "supervisor",                     # ФИО научного руководителя
    "supervisor_department",          # Подразделение МИЭМ (выпадающий список)
    "supervisor_grnti",               # Шифр ГРНТИ
    "supervisor_team",                # Команда (количество, состав, роли)
    "supervisor_competencies",        # Ключевая компетенция
    "supervisor_publications",        # Публикации (до 5)
    "supervisor_grants",              # Гранты (фонд, период, объём)
    "supervisor_niokr",               # Выполненные или идущие НИОКР
    "supervisor_rid_protected",       # Охраняемые РИД
    "supervisor_rid_unprotected",     # Неоформленные РИД (выбрать)
    "supervisor_ugt",                 # УГТ задела научной группы (1-9)
    "supervisor_next_ugt",            # Что нужно для следующего УГТ
    "project_name",
    "customer",
    "problem",
    "competitor",
    "advantage",
    "partner",
    "role_miem",
    "horizon",
    "funding_source",
    "lifecycle_stage",
    "sales_stage",
    "stage_change_reason",
    "stage_change_date"
]

# ---------- ФУНКЦИИ РАБОТЫ С ДАННЫМИ ----------
def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, dtype={"id": int})
    else:
        df = pd.DataFrame(columns=COLUMNS)
        df["id"] = df["id"].astype(int)

    # Добавляем недостающие колонки (миграция)
    for col in COLUMNS:
        if col not in df.columns:
            if col == "id":
                df[col] = 0
            elif col == "supervisor_ugt":
                df[col] = 1
            else:
                df[col] = ""

    # Приводим текстовые колонки к строкам
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

# ---------- СПИСКИ ДЛЯ ВЫБОРА (из файла шаблона) ----------
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

# Списки из листа "Списки" (по данным файла)
UNPROTECTED_RID_LIST = [
    "Алгоритмы/методы", "ПО", "КД", "Прототип", "Ноу-хау",
    "Публикации с тех. описанием", "Технология", "Бизнес-модель", "Другой неоф. РИД"
]
DEPARTMENT_LIST = [
    "Центр квантовых метаматериалов",
    "Международная лаборатория физики элементарных частиц",
    "Научно-технический центр прикладной электроники",
    "Научно-учебная лаборатория квантовой наноэлектроники",
    "Научно-учебная лаборатория телекоммуникационных систем",
    "Научная лаборатория Интернета вещей и киберфизических систем",
    "Учебно-исследовательская лаборатория функциональной безопасности космических аппаратов и систем",
    "Лаборатория «Математические методы естествознания»",
    "Лаборатория вычислительной физики",
    "Лаборатория динамических систем и приложений",
    "Учебно-исследовательская лаборатория Интернет технологий и сервисов",
    "Учебная лаборатория 3Д-визуализации и компьютерной графики",
    "Учебная лаборатория элементов и устройств встраиваемых систем",
    "Учебная лаборатория систем автоматизированного проектирования",
    "Учебная лаборатория сетевых технологий",
    "Учебная лаборатория сетевых видеотехнологий",
    "Учебная лаборатория моделирования систем защиты информации и криптографии",
    "Учебная лаборатория общей и квантовой физики",
    "Учебная лаборатория квантовых технологий",
    "Учебная лаборатория метрологии и измерительных технологий",
    "Учебная лаборатория надежности электронных средств киберфизических систем",
    "Учебная лаборатория моделирования и проектирования электронных компонентов и устройств",
    "Учебная лаборатория микроволновой и оптоэлектронной инженерии",
    "Учебная лаборатория телекоммуникационных технологий и систем связи",
    "Учебная лаборатория электроники и схемотехники",
    "Другое"
]

# ---------- НАВИГАЦИЯ (новый порядок) ----------
st.sidebar.title("Навигация")
page = st.sidebar.radio("Перейти", [
    "👨‍🔬 Научные руководители",
    "📋 Проекты",
    "📊 Дашборд",
    "💾 Экспорт / Импорт"
])

# ---------- СТРАНИЦА "НАУЧНЫЕ РУКОВОДИТЕЛИ" (карточка в порядке шаблона) ----------
if page == "👨‍🔬 Научные руководители":
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
            # Поля в порядке шаблона (A4-A14 + дополнительные)
            # ФИО (уже выбрано)
            st.text_input("ФИО научного руководителя", value=selected_sup, disabled=True)

            # Подразделение МИЭМ (выпадающий список)
            current_dept = sup_df.iloc[0]["supervisor_department"] if "supervisor_department" in sup_df.columns else ""
            dept_idx = DEPARTMENT_LIST.index(current_dept) if current_dept in DEPARTMENT_LIST else 0
            department = st.selectbox("Подразделение МИЭМ (выбрать)", DEPARTMENT_LIST, index=dept_idx)

            # Научное направление (шифр ГРНТИ)
            grnti = st.text_input("Научное направление (шифр ГРНТИ)", value=sup_df.iloc[0]["supervisor_grnti"] if "supervisor_grnti" in sup_df.columns else "")

            # Команда (количество, состав, ФИО)
            team = st.text_area("Команда (количество, состав, ФИО основных штатных участников)", value=sup_df.iloc[0]["supervisor_team"] if "supervisor_team" in sup_df.columns else "")

            # Ключевая компетенция
            comp = st.text_area("Ключевая компетенция", value=sup_df.iloc[0]["supervisor_competencies"] if "supervisor_competencies" in sup_df.columns else "")

            # Публикации (до 5)
            pubs = st.text_area("Публикации (до 5)", value=sup_df.iloc[0]["supervisor_publications"] if "supervisor_publications" in sup_df.columns else "")

            # Гранты (фонд, период, объём)
            grants = st.text_area("Гранты (фонд, период, объём)", value=sup_df.iloc[0]["supervisor_grants"] if "supervisor_grants" in sup_df.columns else "")

            # Выполненные или идущие НИОКР
            niokr = st.text_area("Выполненные или идущие НИОКР", value=sup_df.iloc[0]["supervisor_niokr"] if "supervisor_niokr" in sup_df.columns else "")

            # Охраняемые РИД
            rid_protected = st.text_area("Охраняемые РИД", value=sup_df.iloc[0]["supervisor_rid_protected"] if "supervisor_rid_protected" in sup_df.columns else "")

            # Неоформленные РИД (выпадающий список)
            current_unprot = sup_df.iloc[0]["supervisor_rid_unprotected"] if "supervisor_rid_unprotected" in sup_df.columns else ""
            unprot_idx = UNPROTECTED_RID_LIST.index(current_unprot) if current_unprot in UNPROTECTED_RID_LIST else 0
            rid_unprotected = st.selectbox("Неоформленные РИД (выбрать)", UNPROTECTED_RID_LIST, index=unprot_idx)

            # УГТ задела научной группы (выбрать) – слайдер 1-9
            current_ugt = sup_df.iloc[0]["supervisor_ugt"] if "supervisor_ugt" in sup_df.columns else 1
            ugt = st.slider("УГТ задела научной группы (выбрать)", 1, 9, value=int(current_ugt))

            # Что нужно для следующего УГТ
            next_ugt = st.text_area("Что нужно для следующего УГТ", value=sup_df.iloc[0]["supervisor_next_ugt"] if "supervisor_next_ugt" in sup_df.columns else "")

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
                    st.success("Данные руководителя сохранены")
                    st.rerun()
                else:
                    st.error("Руководитель не найден")

        # Таблица проектов руководителя (русские заголовки)
        st.subheader("Проекты руководителя")
        if not sup_df.empty:
            display_df = sup_df[["project_name", "customer", "lifecycle_stage", "sales_stage", "funding_source"]].copy()
            display_df.columns = ["Название проекта", "Заказчик", "Стадия ЖЦ продукта", "Этап продвижения", "Источник финансирования"]
            st.dataframe(display_df, hide_index=True, use_container_width=True)

        # Прогресс УГТ задела
        current_ugt_val = sup_df.iloc[0]["supervisor_ugt"] if "supervisor_ugt" in sup_df.columns else 1
        st.subheader(f"УГТ задела научной группы: {current_ugt_val}")
        st.progress(current_ugt_val / 9.0)

        # Детализация проектов
        for _, row in sup_df.iterrows():
            with st.expander(f"{row['project_name']} – {row['customer']}"):
                st.write(f"**Решаемая проблема:** {row['problem']}")
                st.write(f"**Конкурент:** {row['competitor']} → **Преимущество:** {row['advantage']}")
                st.write(f"**Партнёр:** {row['partner']} | **Роль МИЭМ:** {row['role_miem']}")
                st.write(f"**Горизонт реализации:** {row['horizon']} | **Источник финансирования:** {row['funding_source']}")
                st.write(f"**Стадия ЖЦ продукта:** {row['lifecycle_stage']} | **Этап продвижения:** {row['sales_stage']}")

# ---------- СТРАНИЦА "ПРОЕКТЫ" (без поля подразделения) ----------
elif page == "📋 Проекты":
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
        display_df = filtered_df[["project_name", "customer", "supervisor", "lifecycle_stage", "sales_stage", "funding_source"]].copy()
        display_df.columns = ["Название проекта", "Заказчик", "Научный руководитель", "Стадия ЖЦ продукта", "Этап продвижения", "Источник финансирования"]
        st.dataframe(display_df, hide_index=True, use_container_width=True)
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
            submitted = st.form_submit_button("Создать")
            if submitted:
                if sup and proj and cust:
                    new_id = get_next_id(df)
                    new_row = pd.DataFrame([{
                        "id": new_id,
                        "supervisor": sup,
                        # значения для руководителя будут заполнены позже в карточке
                        "supervisor_department": "", "supervisor_grnti": "", "supervisor_team": "",
                        "supervisor_competencies": "", "supervisor_publications": "", "supervisor_grants": "",
                        "supervisor_niokr": "", "supervisor_rid_protected": "", "supervisor_rid_unprotected": "",
                        "supervisor_ugt": 1, "supervisor_next_ugt": "",
                        "project_name": proj, "customer": cust, "problem": prob,
                        "competitor": comp, "advantage": adv, "partner": part, "role_miem": role,
                        "horizon": hor, "funding_source": funding,
                        "lifecycle_stage": lc_stage, "sales_stage": sales_stage,
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
                col1, col2 = st.columns(2)
                with col1:
                    saved = st.form_submit_button("Сохранить")
                with col2:
                    deleted = st.form_submit_button("Удалить проект")
                if saved:
                    df.loc[df["id"] == selected_id, ["supervisor", "project_name", "customer", "problem", "competitor",
                                                     "advantage", "partner", "role_miem", "horizon", "funding_source",
                                                     "lifecycle_stage", "sales_stage"]] = \
                        [sup, proj, cust, prob, comp, adv, part, role, hor, funding, lc_stage, sales_stage]
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

# ---------- СТРАНИЦА "ДАШБОРД" (улучшенные графики) ----------
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

        # Воронка продаж (этапы продвижения) с улучшенными осями
        st.subheader("Воронка этапов продвижения")
        sales_counts = df["sales_stage"].value_counts().reindex(SALES_STAGES, fill_value=0).reset_index()
        sales_counts.columns = ["Этап продвижения", "Количество проектов"]
        chart = alt.Chart(sales_counts).mark_bar(color="steelblue").encode(
            x=alt.X("Этап продвижения", sort=SALES_STAGES, title="Этап продвижения", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            y=alt.Y("Количество проектов", title="Количество проектов", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold'))
        ).properties(width=700, height=400)
        # Добавляем пунктирные горизонтальные линии (среднее или просто линии через правило)
        mean_count = sales_counts["Количество проектов"].mean()
        rule = alt.Chart(pd.DataFrame({'y': [mean_count]})).mark_rule(strokeDash=[5,5], color='gray').encode(y='y')
        st.altair_chart(chart + rule, use_container_width=True)

        # Распределение по стадиям ЖЦ
        st.subheader("Стадии жизненного цикла продукта")
        lc_counts = df["lifecycle_stage"].value_counts().reindex(LIFECYCLE_STAGES, fill_value=0).reset_index()
        lc_counts.columns = ["Стадия ЖЦ", "Количество проектов"]
        chart2 = alt.Chart(lc_counts).mark_bar(color="darkorange").encode(
            x=alt.X("Стадия ЖЦ", sort=LIFECYCLE_STAGES, title="Стадия ЖЦ", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold')),
            y=alt.Y("Количество проектов", title="Количество проектов", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold'))
        ).properties(width=700, height=400)
        mean_lc = lc_counts["Количество проектов"].mean()
        rule2 = alt.Chart(pd.DataFrame({'y': [mean_lc]})).mark_rule(strokeDash=[5,5], color='gray').encode(y='y')
        st.altair_chart(chart2 + rule2, use_container_width=True)

        # Распределение по УГТ задела научных групп (целые числа на оси)
        st.subheader("Уровень УГТ задела научных групп")
        ugt_counts = df["supervisor_ugt"].value_counts().sort_index().reset_index()
        ugt_counts.columns = ["УГТ", "Количество руководителей"]
        # Преобразуем УГТ в целочисленный тип для оси
        ugt_counts["УГТ"] = ugt_counts["УГТ"].astype(int)
        chart3 = alt.Chart(ugt_counts).mark_bar(color="green").encode(
            x=alt.X("УГТ:Q", title="УГТ", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold', format='d')),
            y=alt.Y("Количество руководителей:Q", title="Количество", axis=alt.Axis(labelFontSize=12, titleFontSize=14, titleFontWeight='bold'))
        ).properties(width=600, height=400)
        mean_ugt = ugt_counts["Количество руководителей"].mean()
        rule3 = alt.Chart(pd.DataFrame({'y': [mean_ugt]})).mark_rule(strokeDash=[5,5], color='gray').encode(y='y')
        st.altair_chart(chart3 + rule3, use_container_width=True)

        # Последние изменения
        st.subheader("Последние изменения проектов")
        if "stage_change_date" in df.columns and not df["stage_change_date"].isna().all():
            history = df[["project_name", "customer", "sales_stage", "stage_change_date", "stage_change_reason"]].dropna(subset=["stage_change_date"])
            history = history.sort_values("stage_change_date", ascending=False).head(10)
            history.columns = ["Название проекта", "Заказчик", "Этап продвижения", "Дата изменения", "Причина"]
            st.dataframe(history, hide_index=True, use_container_width=True)

# ---------- СТРАНИЦА "ЭКСПОРТ / ИМПОРТ" ----------
elif page == "💾 Экспорт / Импорт":
    st.header("Резервное копирование данных")
    df = load_data()

    st.subheader("📥 Скачать текущую базу")
    if not df.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button(
            label="Скачать projects.xlsx",
            data=output.getvalue(),
            file_name="projects_backup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Нет данных для экспорта")

    st.divider()

    st.subheader("📤 Восстановить из файла")
    uploaded = st.file_uploader("Загрузите ранее сохранённый файл .xlsx", type=["xlsx"])
    if uploaded is not None:
        try:
            backup_df = pd.read_excel(uploaded)
            required_cols = ["id", "supervisor", "project_name", "customer", "supervisor_ugt"]
            if all(col in backup_df.columns for col in required_cols):
                backup_df = backup_df.reset_index(drop=True)
                backup_df.to_excel(DATA_FILE, index=False)
                st.success("База данных восстановлена! Перезагрузите страницу.")
                st.rerun()
            else:
                st.error(f"Файл не содержит необходимых колонок: {required_cols}")
        except Exception as e:
            st.error(f"Ошибка чтения файла: {e}")
