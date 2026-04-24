import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="CRM - Управление проектами", layout="wide")
st.title("📋 Управление проектами (CRM)")
DATA_FILE = "projects.xlsx"

def load_data():
    """Загружает данные из Excel, фильтрует мусор, добавляет недостающие колонки."""
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, dtype={"id": int})
    else:
        df = pd.DataFrame(columns=[
            "id", "name", "organization", "department", "stage", "ugt",
            "stage_change_reason", "stage_change_date"
        ])
        df["id"] = df["id"].astype(int)

    # Удаляем записи с организацией "[вручную]" (пункт 4)
    df = df[~df["organization"].astype(str).str.contains(r"\[вручную\]", na=False, case=False)]
    df = df.reset_index(drop=True)

    # Добавляем недостающие колонки
    if "stage" not in df.columns:
        df["stage"] = "Квалификация"
    if "ugt" not in df.columns:
        df["ugt"] = 1
    if "stage_change_reason" not in df.columns:
        df["stage_change_reason"] = ""
    if "stage_change_date" not in df.columns:
        df["stage_change_date"] = pd.NaT
    if "department" not in df.columns:
        df["department"] = ""

    # Приводим типы
    df["ugt"] = pd.to_numeric(df["ugt"], errors="coerce").fillna(1).astype(int)
    df["stage_change_date"] = pd.to_datetime(df["stage_change_date"], errors="coerce")
    if not df.empty:
        df["id"] = df["id"].astype(int)
        df = df.sort_values("id").reset_index(drop=True)
    return df

def save_data(df):
    """Сохраняет DataFrame в Excel."""
    df.to_excel(DATA_FILE, index=False)

def get_next_id(df):
    """Возвращает следующий доступный ID."""
    return int(df["id"].max() + 1) if not df.empty else 1# --- Боковая навигация ---
st.sidebar.title("Навигация")
page = st.sidebar.radio("Перейти", [
    "Дашборд", "Паспорта (проекты)", "Контрагенты", "Совместная деятельность", "Импорт из Excel"
])

if page == "Паспорта (проекты)":
    st.header("📌 Паспорта проектов")
    df = load_data()

    # --- Список проектов с нумерацией с 1 (пункт 3) ---
    st.subheader("Список проектов")
    if not df.empty:
        display_df = df.copy()
        display_df.insert(0, "№", range(1, len(display_df)+1))
        st.dataframe(
            display_df[["№", "name", "organization", "stage", "ugt"]],
            hide_index=True,
            use_container_width=True
        )
    else:
        st.info("Нет проектов. Добавьте первый через форму создания или импорт.")

    # --- Выбор проекта для редактирования ---
    st.subheader("Редактирование проекта")
    ids = df["id"].tolist() if not df.empty else []
    if ids:
        selected_id = st.selectbox("Выберите ID проекта", ids, format_func=lambda x: f"ID {x}")
    else:
        selected_id = None
        st.warning("Нет проектов для редактирования.")

    if selected_id is not None:
        project = df[df["id"] == selected_id].iloc[0]
        original_stage = project["stage"]

        with st.form(key="edit_form"):
            st.subheader(f"✏️ Редактирование проекта ID {selected_id}")

            name = st.text_input("Название проекта", value=project["name"])
            organization = st.text_input("Организация", value=project["organization"])

            # --- Подразделение (безопасный выбор, пункт 2) ---
            existing_depts = sorted(df["department"].dropna().unique())
            if existing_depts:
                dept_options = ["(Выберите или введите новое)"] + existing_depts
                if project["department"] in existing_depts:
                    default_index = existing_depts.index(project["department"]) + 1
                else:
                    default_index = 0
                choice = st.selectbox("Подразделение", dept_options, index=default_index)
                if choice == "(Выберите или введите новое)":
                    department = st.text_input("Новое подразделение", value=project["department"] if project["department"] not in existing_depts else "")
                else:
                    department = choice
            else:
                department = st.text_input("Подразделение", value=project["department"])

            # --- Этапы сделки (6 стадий) ---
            stage_options = ["Квалификация", "Формирование решения", "Переговоры", "Закрытие", "Внедрён / Завершён", "Отклонён"]
            stage_idx = stage_options.index(project["stage"]) if project["stage"] in stage_options else 0
            new_stage = st.selectbox("Этап сделки", stage_options, index=stage_idx)

            # --- Уровень готовности технологии (УГТ) ---
            ugt = st.number_input("УГТ (1–9)", min_value=1, max_value=9, step=1, value=int(project["ugt"]))

            # --- Цифровой след: отображение последней причины ---
            last_reason = project["stage_change_reason"] if pd.notna(project["stage_change_reason"]) else ""
            st.text_area("Последняя причина смены этапа (история)", value=last_reason, disabled=True, help="Заполняется автоматически при смене этапа")

            change_reason = ""
            if new_stage != original_stage:
                change_reason = st.text_area(f"✏️ Причина перехода на этап «{new_stage}»", placeholder="Обязательно укажите причину", key="reason_input")

            col1, col2 = st.columns(2)
            submitted = col1.form_submit_button("💾 Сохранить")
            delete = col2.form_submit_button("🗑️ Удалить проект")

            if submitted:
                if new_stage != original_stage and not change_reason.strip():
                    st.error("При смене этапа необходимо указать причину перехода.")
                    st.stop()

                # Сохраняем основные поля
                df.loc[df["id"] == selected_id, "name"] = name
                df.loc[df["id"] == selected_id, "organization"] = organization
                df.loc[df["id"] == selected_id, "department"] = department
                df.loc[df["id"] == selected_id, "ugt"] = ugt

                if new_stage != original_stage:
                    df.loc[df["id"] == selected_id, "stage"] = new_stage
                    df.loc[df["id"] == selected_id, "stage_change_reason"] = change_reason
                    df.loc[df["id"] == selected_id, "stage_change_date"] = datetime.now()

                save_data(df)
                st.success("Изменения сохранены!")
                st.rerun()

            if delete:
                df = df[df["id"] != selected_id]
                save_data(df)
                st.success("Проект удалён!")
                st.rerun()

    # --- Создание нового проекта ---
    st.subheader("➕ Создать новый проект")
    with st.form(key="new_form"):
        new_name = st.text_input("Название проекта*")
        new_org = st.text_input("Организация*")
        new_dept = st.text_input("Подразделение (опционально)")
        new_stage = st.selectbox("Начальный этап", stage_options, index=0)
        new_ugt = st.number_input("УГТ", min_value=1, max_value=9, value=1, step=1)
        create = st.form_submit_button("Создать")
        if create:
            if new_name and new_org:
                new_id = get_next_id(df)
                new_row = pd.DataFrame([{
                    "id": new_id,
                    "name": new_name,
                    "organization": new_org,
                    "department": new_dept,
                    "stage": new_stage,
                    "ugt": new_ugt,
                    "stage_change_reason": "Создание проекта",
                    "stage_change_date": datetime.now()
                }])
                df = pd.concat([df, new_row], ignore_index=True)
                save_data(df)
                st.success(f"Проект ID {new_id} создан!")
                st.rerun()
            else:
                st.error("Название и организация обязательны")

elif page == "Дашборд":
    st.header("📊 Дашборд")
    df = load_data()
    if not df.empty:
        col1, col2, col3 = st.columns(3)
        col1.metric("Всего проектов", len(df))
        active = len(df[~df["stage"].isin(["Внедрён / Завершён", "Отклонён"])])
        col2.metric("Активных проектов", active)
        col3.metric("Средний УГТ", round(df["ugt"].mean(), 1))

        st.subheader("Воронка по этапам")
        stage_order = ["Квалификация", "Формирование решения", "Переговоры", "Закрытие", "Внедрён / Завершён", "Отклонён"]
        stage_counts = df["stage"].value_counts().reindex(stage_order, fill_value=0)
        st.bar_chart(stage_counts)

        st.subheader("📜 Цифровой след (последние изменения этапов)")
        history = df[["name", "stage", "stage_change_date", "stage_change_reason"]].dropna(subset=["stage_change_date"])
        history = history.sort_values("stage_change_date", ascending=False).head(10)
        st.dataframe(history, hide_index=True, use_container_width=True)
    else:
        st.info("Нет данных для отображения.")elif page == "Контрагенты":
    st.header("🏢 Контрагенты")
    st.info("Страница в разработке. Здесь будет список организаций.")

elif page == "Совместная деятельность":
    st.header("🤝 Совместная деятельность")
    st.info("Страница в разработке.")

elif page == "Импорт из Excel":
    st.header("📂 Импорт из Excel")
    uploaded = st.file_uploader("Загрузите Excel-файл с проектами", type=["xlsx"])
    if uploaded:
        try:
            new_df = pd.read_excel(uploaded)
            required = ["name", "organization"]
            if all(col in new_df.columns for col in required):
                current_df = load_data()
                next_id = get_next_id(current_df)
                imported = 0
                for idx, row in new_df.iterrows():
                    # Пропускаем записи с [вручную]
                    if "[вручную]" in str(row["organization"]):
                        continue
                    new_row = pd.DataFrame([{
                        "id": next_id + idx,
                        "name": row["name"],
                        "organization": row["organization"],
                        "department": row.get("department", ""),
                        "stage": row.get("stage", "Квалификация"),
                        "ugt": row.get("ugt", 1),
                        "stage_change_reason": "Импорт из Excel",
                        "stage_change_date": datetime.now()
                    }])
                    current_df = pd.concat([current_df, new_row], ignore_index=True)
                    imported += 1
                save_data(current_df)
                st.success(f"Импортировано {imported} проектов (пропущены с '[вручную]')")
                st.rerun()
            else:
                st.error("Файл должен содержать колонки: name, organization")
        except Exception as e:
            st.error(f"Ошибка при импорте: {e}")

# Конец файла (Streamlit сам запускает)
