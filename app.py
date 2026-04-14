import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String, Text, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from datetime import date
import re

# ----- Настройка базы данных -----
engine = create_engine('sqlite:///crm.db', echo=False)
Base = declarative_base()
Session = sessionmaker(bind=engine)

# ----- Модели данных -----

class Zadel(Base):
    __tablename__ = 'zadels'
    id = Column(Integer, primary_key=True)
    name = Column(String(200), nullable=False)
    grnti = Column(String(50))
    collective = Column(String(100))
    team = Column(Text)
    competencies = Column(Text)
    publications = Column(Text)
    grants = Column(Text)
    niokr = Column(Text)
    rid_protected = Column(Text)
    rid_unprotected = Column(String(100))
    ugt = Column(Integer)
    next_ugt_needs = Column(Text)

class Organization(Base):
    __tablename__ = 'organizations'
    id = Column(Integer, primary_key=True)
    name = Column(String(200), nullable=False)
    industry = Column(String(100))
    org_type = Column(String(50))
    size = Column(String(50))
    priority = Column(Integer, default=5)
    notes = Column(Text)

class Contact(Base):
    __tablename__ = 'contacts'
    id = Column(Integer, primary_key=True)
    org_id = Column(Integer, ForeignKey('organizations.id'), nullable=False)
    name = Column(String(200), nullable=False)
    position = Column(String(100))
    email = Column(String(100))
    phone = Column(String(50))
    organization = relationship("Organization", backref="contacts")

class Barrier(Base):
    __tablename__ = 'barriers'
    id = Column(Integer, primary_key=True)
    name = Column(String(200), nullable=False, unique=True)
    weight = Column(Integer, default=5)

class MiemRole(Base):
    __tablename__ = 'miem_roles'
    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False, unique=True)

class Hypothesis(Base):
    __tablename__ = 'hypotheses'
    id = Column(Integer, primary_key=True)
    name = Column(String(200), nullable=False)
    org_id = Column(Integer, ForeignKey('organizations.id'), nullable=False)
    participation_type = Column(String(50))
    miem_role_id = Column(Integer, ForeignKey('miem_roles.id'))
    zadel_id = Column(Integer, ForeignKey('zadels.id'))
    ugt = Column(Integer)
    competitor = Column(String(200))
    advantage = Column(Text)
    budget = Column(Integer)
    horizon = Column(String(50))
    status = Column(String(50))
    responsible = Column(String(100))
    docs_link = Column(String(500))
    created_at = Column(String(20))
    organization = relationship("Organization", backref="hypotheses")
    miem_role = relationship("MiemRole")
    zadel = relationship("Zadel")

class HypothesisBarrier(Base):
    __tablename__ = 'hypothesis_barriers'
    id = Column(Integer, primary_key=True)
    hypothesis_id = Column(Integer, ForeignKey('hypotheses.id'))
    barrier_id = Column(Integer, ForeignKey('barriers.id'))

class Task(Base):
    __tablename__ = 'tasks'
    id = Column(Integer, primary_key=True)
    hypothesis_id = Column(Integer, ForeignKey('hypotheses.id'), nullable=False)
    name = Column(String(200), nullable=False)
    responsible = Column(String(100))
    due_date = Column(String(20))
    status = Column(String(50), default='не выполнена')
    hypothesis = relationship("Hypothesis", backref="tasks")

# Создаём таблицы
Base.metadata.create_all(engine)

# Инициализация справочников
def init_dicts():
    session = Session()
    barriers_with_weights = [
        ("Нет оформленных прав на РИД", 8),
        ("Нет прототипа в реальных условиях", 7),
        ("Нет подходящего партнёра", 6),
        ("Нет времени", 5),
        ("Нет исполнителей", 7),
        ("Нет понимания рынка", 9),
        ("Нет коммерческого потенциала", 10),
        ("Отсутствие инфраструктуры", 6),
        ("Другой барьер", 5)
    ]
    for name, weight in barriers_with_weights:
        b = session.query(Barrier).filter(Barrier.name == name).first()
        if not b:
            session.add(Barrier(name=name, weight=weight))
        else:
            b.weight = weight
    roles_list = ["Субподрядчик", "Соисполнитель", "Лицензиар", "Сервисный центр",
                  "Технологический аудитор", "Консультант", "Другая роль"]
    for r in roles_list:
        if not session.query(MiemRole).filter(MiemRole.name == r).first():
            session.add(MiemRole(name=r))
    session.commit()
    session.close()

init_dicts()

# ----- Вспомогательные функции -----
def get_org_dict():
    session = Session()
    orgs = session.query(Organization).all()
    res = {o.id: o.name for o in orgs}
    session.close()
    return res

def get_zadel_dict():
    session = Session()
    zadels = session.query(Zadel).all()
    res = {z.id: z.name for z in zadels}
    session.close()
    return res

def get_role_dict():
    session = Session()
    roles = session.query(MiemRole).all()
    res = {r.id: r.name for r in roles}
    session.close()
    return res

# ----- Интерфейс -----
st.set_page_config(page_title="CRM МИЭМ", layout="wide")

st.sidebar.title("Навигация")
page = st.sidebar.radio("Перейти", ["Дашборд", "Паспорта (заделы)", "Контрагенты", "Совместная деятельность", "Импорт из Excel"])

# -------------------- ДАШБОРД --------------------
if page == "Дашборд":
    st.title("Управление коммерциализацией МИЭМ")
    st.header("Дашборд")
    
    session = Session()
    
    # ----- Ключевые метрики -----
    total_hypotheses = session.query(Hypothesis).count()
    active_hypotheses = session.query(Hypothesis).filter(Hypothesis.status.in_(["идентифицирована", "квалифицирована", "предложение отправлено", "переговоры"])).count()
    contracts = session.query(Hypothesis).filter(Hypothesis.status == "контракт").count()
    avg_ugt = session.query(Zadel.ugt).filter(Zadel.ugt.isnot(None)).all()
    avg_ugt_value = sum([a[0] for a in avg_ugt]) / len(avg_ugt) if avg_ugt else 0
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Всего гипотез", total_hypotheses)
    col2.metric("Активных", active_hypotheses)
    col3.metric("Контрактов", contracts)
    col4.metric("Средний УГТ", f"{avg_ugt_value:.1f}")
    
    st.divider()
    
    # ----- Матрица "Организация × Коллектив" с фильтрами -----
    st.subheader("Матрица «Организация × Коллектив»")
    
    # Фильтры
    col_filter1, col_filter2, col_filter3 = st.columns(3)
    with col_filter1:
        participation_filter = st.selectbox("Тип участия", ["Все", "Заказчик", "Партнёр"])
    with col_filter2:
        ugt_min = st.number_input("Минимальный УГТ", min_value=1, max_value=9, value=1)
    with col_filter3:
        role_filter = st.selectbox("Роль МИЭМ", ["Все"] + [r.name for r in session.query(MiemRole).all()])
    
    # Получаем организации и коллективы
    orgs = session.query(Organization).all()
    collectives = session.query(Zadel.collective).distinct().filter(Zadel.collective.isnot(None)).all()
    collectives = [c[0] for c in collectives if c[0]]
    
    if orgs and collectives:
        matrix_data = []
        for org in orgs:
            row = {"Организация": org.name}
            for coll in collectives:
                # Базовый запрос гипотез для пары организация-коллектив
                query = session.query(Hypothesis).join(Zadel, Hypothesis.zadel_id == Zadel.id).filter(
                    Hypothesis.org_id == org.id,
                    Zadel.collective == coll
                )
                # Применяем фильтры
                if participation_filter != "Все":
                    query = query.filter(Hypothesis.participation_type == participation_filter)
                if role_filter != "Все":
                    role = session.query(MiemRole).filter(MiemRole.name == role_filter).first()
                    if role:
                        query = query.filter(Hypothesis.miem_role_id == role.id)
                # Фильтр по УГТ задела
                query = query.filter(Zadel.ugt >= ugt_min)
                count = query.count()
                row[coll] = count
            matrix_data.append(row)
        df_matrix = pd.DataFrame(matrix_data)
        st.dataframe(df_matrix, use_container_width=True)
    else:
        st.info("Недостаточно данных для построения матрицы. Добавьте организации и заделы.")
    
    st.divider()
    
    # ----- Топ‑5 гипотез по приоритету (с учётом весов барьеров) -----
    st.subheader("Топ‑5 приоритетных гипотез")
    
    def calculate_priority(hypothesis):
        org_priority = hypothesis.organization.priority if hypothesis.organization else 5
        ugt = hypothesis.ugt if hypothesis.ugt else 1
        # Сумма весов барьеров
        barriers = session.query(HypothesisBarrier).filter(HypothesisBarrier.hypothesis_id == hypothesis.id).all()
        total_weight = 0
        for hb in barriers:
            b = session.query(Barrier).filter(Barrier.id == hb.barrier_id).first()
            if b:
                total_weight += b.weight
        # Штраф: сумма весов / 10 (нормировка, чтобы максимальный штраф ~ 10)
        penalty = total_weight / 10
        # Приоритет: (приоритет организации * 0.5 + УГТ * 0.5) - штраф
        priority = (org_priority * 0.5 + ugt * 0.5) - penalty
        return max(priority, 0)
    
    hypotheses = session.query(Hypothesis).all()
    if hypotheses:
        hyp_list = []
        for h in hypotheses:
            hyp_list.append({
                "id": h.id,
                "name": h.name,
                "organization": h.organization.name if h.organization else "—",
                "status": h.status,
                "priority": calculate_priority(h)
            })
        df_hyp = pd.DataFrame(hyp_list)
        df_hyp = df_hyp.sort_values("priority", ascending=False).head(5)
        
        for _, row in df_hyp.iterrows():
            with st.container():
                st.markdown(f"**{row['name']}**  (приоритет: {row['priority']:.1f})")
                st.write(f"Организация: {row['organization']} | Статус: {row['status']}")
                # Кнопка для быстрого перехода к редактированию (позже)
                if st.button(f"Перейти к гипотезе", key=f"goto_{row['id']}"):
                    st.session_state['edit_hypothesis_id'] = row['id']
                    st.switch_page("app.py")  # не работает в Streamlit, лучше установить сессию и редирект
                    # Временно просто сообщаем
                    st.info(f"Вы выбрали гипотезу ID {row['id']}. Перейдите на вкладку «Совместная деятельность» для редактирования.")
                st.divider()
    else:
        st.info("Нет гипотез для отображения.")
    
    st.divider()
    
    # ----- Виджет авторекомендаций с весами барьеров -----
    st.subheader("Рекомендации")
    
    recommendations = []
    
    # 1. Заделы с УГТ ≤ 3 без активных гипотез
    low_ugt_zadels = session.query(Zadel).filter(Zadel.ugt <= 3).all()
    for z in low_ugt_zadels:
        hyp_count = session.query(Hypothesis).filter(Hypothesis.zadel_id == z.id).count()
        if hyp_count == 0:
            recommendations.append(f"📌 **Задел «{z.name}»** (УГТ={z.ugt}) не используется. Рекомендуется подать заявку на грант или найти индустриального партнёра для повышения УГТ.")
    
    # 2. Неоформленные РИД = "Алгоритмы/методы" без патента
    unprotected_zadels = session.query(Zadel).filter(Zadel.rid_unprotected == "Алгоритмы/методы").all()
    for z in unprotected_zadels:
        if not z.rid_protected:
            recommendations.append(f"🔐 **Задел «{z.name}»**: неоформленные РИД – «Алгоритмы/методы». Оформите патент или свидетельство о ПО.")
    
    # 3. Организации с высоким приоритетом (≥8) без активных гипотез
    high_priority_orgs = session.query(Organization).filter(Organization.priority >= 8).all()
    for org in high_priority_orgs:
        hyp_count = session.query(Hypothesis).filter(Hypothesis.org_id == org.id).count()
        if hyp_count == 0:
            recommendations.append(f"🏢 **Организация «{org.name}»** имеет высокий приоритет ({org.priority}), но нет активных гипотез. Инициируйте контакт.")
    
    # 4. Гипотезы с барьером "Нет подходящего партнёра" (вес 6)
    barrier_partner = session.query(Barrier).filter(Barrier.name == "Нет подходящего партнёра").first()
    if barrier_partner:
        hyp_with_partner_barrier = session.query(HypothesisBarrier).filter(HypothesisBarrier.barrier_id == barrier_partner.id).all()
        for hb in hyp_with_partner_barrier:
            hyp = session.query(Hypothesis).filter(Hypothesis.id == hb.hypothesis_id).first()
            if hyp:
                recommendations.append(f"🤝 **Гипотеза «{hyp.name}»**: барьер «Нет подходящего партнёра». Рекомендуется найти партнёра в индустрии {hyp.organization.industry if hyp.organization else '?'}.")
    
    # 5. Заделы с УГТ ≥ 6 без гипотез с ролью Лицензиар
    high_ugt_zadels = session.query(Zadel).filter(Zadel.ugt >= 6).all()
    lic_role = session.query(MiemRole).filter(MiemRole.name == "Лицензиар").first()
    if lic_role:
        for z in high_ugt_zadels:
            hyp_count = session.query(Hypothesis).filter(Hypothesis.zadel_id == z.id, Hypothesis.miem_role_id == lic_role.id).count()
            if hyp_count == 0:
                recommendations.append(f"💼 **Задел «{z.name}»** (УГТ={z.ugt}) готов к лицензированию. Инициируйте предложение роли «Лицензиар».")
    
    # 6. Гипотезы с высоким суммарным весом барьеров (штраф > 3)
    for h in hypotheses:
        barriers = session.query(HypothesisBarrier).filter(HypothesisBarrier.hypothesis_id == h.id).all()
        total_weight = sum([session.query(Barrier).get(b.barrier_id).weight for b in barriers if session.query(Barrier).get(b.barrier_id)])
        if total_weight > 3:
            recommendations.append(f"⚠️ **Гипотеза «{h.name}»**: суммарный вес барьеров = {total_weight}. Рекомендуется провести анализ и снизить барьеры.")
    
    if recommendations:
        for rec in recommendations[:5]:
            st.info(rec)
    else:
        st.success("Нет активных рекомендаций. Хорошая работа!")
    
    session.close()

# -------------------- ПАСПОРТА (ЗАДЕЛЫ) --------------------
elif page == "Паспорта (заделы)":
    st.header("Паспорта научных заделов (направлений)")
    
    # Форма добавления
    with st.expander("➕ Добавить новый задел", expanded=False):
        with st.form("new_zadel_form"):
            name = st.text_input("Название задела*")
            grnti = st.text_input("Шифр ГРНТИ")
            collective = st.text_input("Коллектив (научная группа)")
            team = st.text_area("Команда (количество, состав, ФИО)")
            competencies = st.text_area("Ключевые компетенции")
            publications = st.text_area("Публикации (до 5)")
            grants = st.text_area("Гранты")
            niokr = st.text_area("Выполненные или идущие НИОКР")
            rid_protected = st.text_area("Охраняемые РИД")
            rid_unprotected = st.selectbox("Неоформленные РИД", ["", "Ноу-хау", "Алгоритмы/методы", "ПО", "Технология", "Бизнес-модель"])
            ugt = st.selectbox("УГТ (уровень готовности технологии)", list(range(1,10)), index=0)
            next_ugt_needs = st.text_area("Что нужно для следующего УГТ")
            submitted = st.form_submit_button("Сохранить задел")
            if submitted and name:
                session = Session()
                new_z = Zadel(name=name, grnti=grnti, collective=collective, team=team,
                              competencies=competencies, publications=publications, grants=grants,
                              niokr=niokr, rid_protected=rid_protected,
                              rid_unprotected=rid_unprotected if rid_unprotected else None,
                              ugt=ugt, next_ugt_needs=next_ugt_needs)
                session.add(new_z)
                session.commit()
                session.close()
                st.success(f"Задел '{name}' добавлен!")
                st.rerun()
            elif submitted and not name:
                st.error("Название обязательно")
    
    # Список заделов с возможностью редактирования и удаления
    st.subheader("Существующие заделы")
    session = Session()
    zadels = session.query(Zadel).all()
    if zadels:
        df = pd.DataFrame([{"ID": z.id, "Название": z.name, "УГТ": z.ugt, "Коллектив": z.collective, "ГРНТИ": z.grnti} for z in zadels])
        st.dataframe(df, use_container_width=True)
        
        selected_id = st.selectbox("Выберите ID задела для просмотра/редактирования", [z.id for z in zadels])
        if selected_id:
            z = session.query(Zadel).filter(Zadel.id == selected_id).first()
            if z:
                with st.expander("Редактировать задел"):
                    with st.form("edit_zadel"):
                        new_name = st.text_input("Название", value=z.name)
                        new_grnti = st.text_input("Шифр ГРНТИ", value=z.grnti or "")
                        new_collective = st.text_input("Коллектив", value=z.collective or "")
                        new_team = st.text_area("Команда", value=z.team or "")
                        new_competencies = st.text_area("Компетенции", value=z.competencies or "")
                        new_publications = st.text_area("Публикации", value=z.publications or "")
                        new_grants = st.text_area("Гранты", value=z.grants or "")
                        new_niokr = st.text_area("НИОКР", value=z.niokr or "")
                        new_rid_protected = st.text_area("Охраняемые РИД", value=z.rid_protected or "")
                        new_rid_unprotected = st.selectbox("Неоформленные РИД", ["", "Ноу-хау", "Алгоритмы/методы", "ПО", "Технология", "Бизнес-модель"], index=["", "Ноу-хау", "Алгоритмы/методы", "ПО", "Технология", "Бизнес-модель"].index(z.rid_unprotected) if z.rid_unprotected else 0)
                        new_ugt = st.selectbox("УГТ", list(range(1,10)), index=z.ugt-1)
                        new_next = st.text_area("Что нужно для следующего УГТ", value=z.next_ugt_needs or "")
                        col_edit1, col_edit2 = st.columns(2)
                        with col_edit1:
                            save = st.form_submit_button("Сохранить изменения")
                        with col_edit2:
                            delete = st.form_submit_button("Удалить задел")
                        if save:
                            z.name = new_name
                            z.grnti = new_grnti
                            z.collective = new_collective
                            z.team = new_team
                            z.competencies = new_competencies
                            z.publications = new_publications
                            z.grants = new_grants
                            z.niokr = new_niokr
                            z.rid_protected = new_rid_protected
                            z.rid_unprotected = new_rid_unprotected if new_rid_unprotected else None
                            z.ugt = new_ugt
                            z.next_ugt_needs = new_next
                            session.commit()
                            st.success("Изменения сохранены")
                            st.rerun()
                        if delete:
                            session.delete(z)
                            session.commit()
                            st.success("Задел удалён")
                            st.rerun()
                st.subheader("Детали задела")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"**Шифр ГРНТИ:** {z.grnti or ''}")
                    st.markdown(f"**Коллектив:** {z.collective or ''}")
                    st.markdown(f"**Команда:** {z.team or ''}")
                    st.markdown(f"**Ключевые компетенции:** {z.competencies or ''}")
                    st.markdown(f"**УГТ:** {z.ugt}")
                with col2:
                    st.markdown(f"**Охраняемые РИД:** {z.rid_protected or ''}")
                    st.markdown(f"**Неоформленные РИД:** {z.rid_unprotected or ''}")
                    st.markdown(f"**Что нужно для следующего УГТ:** {z.next_ugt_needs or ''}")
                with st.expander("Публикации, гранты, НИОКР"):
                    st.markdown(f"**Публикации:**\n{z.publications or '—'}")
                    st.markdown(f"**Гранты:**\n{z.grants or '—'}")
                    st.markdown(f"**НИОКР:**\n{z.niokr or '—'}")
    else:
        st.info("Нет заделов. Добавьте первый через форму.")
    session.close()

# -------------------- КОНТРАГЕНТЫ --------------------
elif page == "Контрагенты":
    st.header("Контрагенты")
    tab1, tab2, tab3 = st.tabs(["Организации", "Контакты", "Потребности"])
    
    # ----- Организации -----
    with tab1:
        st.subheader("Организации")
        with st.expander("➕ Добавить организацию"):
            with st.form("new_org_form"):
                name = st.text_input("Название организации*")
                industry = st.selectbox("Индустрия", ["", "микроэлектроника", "нефтегаз", "IT", "оборонка", "другое"])
                org_type = st.selectbox("Тип заказчика", ["", "государственный", "частный корпоративный", "средний бизнес", "стартап"])
                size = st.selectbox("Размер", ["", "малый", "средний", "крупный"])
                priority = st.slider("Приоритет для МИЭМ (1-10)", 1, 10, 5)
                notes = st.text_area("Заметки / потребности")
                submitted = st.form_submit_button("Сохранить")
                if submitted and name:
                    session = Session()
                    org = Organization(name=name, industry=industry, org_type=org_type, size=size, priority=priority, notes=notes)
                    session.add(org)
                    session.commit()
                    session.close()
                    st.success(f"Организация '{name}' добавлена")
                    st.rerun()
                elif submitted and not name:
                    st.error("Название обязательно")
        
        session = Session()
        orgs = session.query(Organization).all()
        if orgs:
            df_org = pd.DataFrame([{"ID": o.id, "Название": o.name, "Индустрия": o.industry or "", "Тип": o.org_type or "", "Размер": o.size or "", "Приоритет": o.priority} for o in orgs])
            st.dataframe(df_org, use_container_width=True)
            
            selected_org_id = st.selectbox("Выберите организацию для редактирования", [o.id for o in orgs], key="org_select")
            if selected_org_id:
                org = session.query(Organization).filter(Organization.id == selected_org_id).first()
                with st.expander("Редактировать организацию"):
                    with st.form("edit_org"):
                        new_name = st.text_input("Название", value=org.name)
                        new_industry = st.selectbox("Индустрия", ["", "микроэлектроника", "нефтегаз", "IT", "оборонка", "другое"], index=["", "микроэлектроника", "нефтегаз", "IT", "оборонка", "другое"].index(org.industry) if org.industry else 0)
                        new_type = st.selectbox("Тип заказчика", ["", "государственный", "частный корпоративный", "средний бизнес", "стартап"], index=["", "государственный", "частный корпоративный", "средний бизнес", "стартап"].index(org.org_type) if org.org_type else 0)
                        new_size = st.selectbox("Размер", ["", "малый", "средний", "крупный"], index=["", "малый", "средний", "крупный"].index(org.size) if org.size else 0)
                        new_priority = st.slider("Приоритет", 1, 10, value=org.priority)
                        new_notes = st.text_area("Заметки", value=org.notes or "")
                        col1, col2 = st.columns(2)
                        with col1:
                            save = st.form_submit_button("Сохранить")
                        with col2:
                            delete = st.form_submit_button("Удалить организацию")
                        if save:
                            org.name = new_name
                            org.industry = new_industry
                            org.org_type = new_type
                            org.size = new_size
                            org.priority = new_priority
                            org.notes = new_notes
                            session.commit()
                            st.success("Сохранено")
                            st.rerun()
                        if delete:
                            session.delete(org)
                            session.commit()
                            st.success("Организация удалена")
                            st.rerun()
        else:
            st.info("Нет организаций. Добавьте первую.")
        session.close()
    
    # ----- Контакты -----
    with tab2:
        st.subheader("Контакты")
        with st.expander("➕ Добавить контакт"):
            with st.form("new_contact_form"):
                session = Session()
                org_list = session.query(Organization).all()
                org_options = {f"{o.name} (ID {o.id})": o.id for o in org_list}
                session.close()
                if org_options:
                    org_choice = st.selectbox("Организация", list(org_options.keys()))
                    org_id = org_options[org_choice]
                    contact_name = st.text_input("ФИО*")
                    position = st.text_input("Должность")
                    email = st.text_input("Email")
                    phone = st.text_input("Телефон")
                    submitted_contact = st.form_submit_button("Сохранить контакт")
                    if submitted_contact and contact_name:
                        session = Session()
                        new_contact = Contact(org_id=org_id, name=contact_name, position=position, email=email, phone=phone)
                        session.add(new_contact)
                        session.commit()
                        session.close()
                        st.success(f"Контакт {contact_name} добавлен")
                        st.rerun()
                    elif submitted_contact and not contact_name:
                        st.error("ФИО обязательно")
                else:
                    st.warning("Сначала добавьте хотя бы одну организацию")
        
        session = Session()
        contacts = session.query(Contact).all()
        if contacts:
            org_dict = {o.id: o.name for o in session.query(Organization).all()}
            df_contact = pd.DataFrame([{"ID": c.id, "Организация": org_dict.get(c.org_id, "—"), "ФИО": c.name, "Должность": c.position or "", "Email": c.email or "", "Телефон": c.phone or ""} for c in contacts])
            st.dataframe(df_contact, use_container_width=True)
            
            selected_contact_id = st.selectbox("Выберите контакт для удаления", [c.id for c in contacts], key="contact_select")
            if st.button("Удалить выбранный контакт"):
                session.query(Contact).filter(Contact.id == selected_contact_id).delete()
                session.commit()
                st.success("Контакт удалён")
                st.rerun()
        else:
            st.info("Нет контактов. Добавьте первый.")
        session.close()
    
    # ----- Потребности -----
    with tab3:
        st.subheader("Потребности организаций")
        session = Session()
        orgs_need = session.query(Organization).all()
        if orgs_need:
            for o in orgs_need:
                with st.expander(f"{o.name}"):
                    new_note = st.text_area(f"Редактировать потребности / заметки", value=o.notes or "", key=f"note_{o.id}")
                    if st.button(f"Сохранить для {o.name}", key=f"save_{o.id}"):
                        o.notes = new_note
                        session.commit()
                        st.success("Сохранено")
                        st.rerun()
        else:
            st.info("Нет организаций")
        session.close()

# -------------------- СОВМЕСТНАЯ ДЕЯТЕЛЬНОСТЬ --------------------
elif page == "Совместная деятельность":
    st.header("Совместная деятельность")
    tab_hyp, tab_tasks = st.tabs(["Гипотезы и проекты", "Мои задачи"])
    
    with tab_hyp:
        with st.expander("➕ Создать новую гипотезу/проект"):
            with st.form("new_hypothesis"):
                name = st.text_input("Название гипотезы*")
                session = Session()
                orgs = session.query(Organization).all()
                org_options = {o.name: o.id for o in orgs}
                zadels = session.query(Zadel).all()
                zadel_options = {z.name: z.id for z in zadels}
                roles = session.query(MiemRole).all()
                role_options = {r.name: r.id for r in roles}
                session.close()
                if not org_options or not zadel_options:
                    st.warning("Сначала добавьте организации и заделы на соответствующих страницах.")
                    org_id = None
                    zadel_id = None
                else:
                    org_name = st.selectbox("Организация", list(org_options.keys()))
                    org_id = org_options[org_name]
                    zadel_name = st.selectbox("Задел", list(zadel_options.keys()))
                    zadel_id = zadel_options[zadel_name]
                    session2 = Session()
                    ugt_val = session2.query(Zadel).filter(Zadel.id == zadel_id).first().ugt
                    session2.close()
                    st.info(f"УГТ задела: {ugt_val}")
                part_type = st.radio("Тип участия организации", ["Заказчик", "Партнёр"], horizontal=True)
                role_name = st.selectbox("Роль МИЭМ", list(role_options.keys()) if role_options else ["Нет ролей"])
                competitor = st.text_input("Конкурент (название)")
                advantage = st.text_area("Преимущество перед конкурентом")
                session = Session()
                barriers = session.query(Barrier).all()
                barrier_names = [b.name for b in barriers]
                session.close()
                selected_barriers = st.multiselect("Барьеры", barrier_names)
                budget = st.number_input("Оценочный бюджет (руб.)", min_value=0, step=100000, value=0)
                horizon = st.selectbox("Горизонт реализации", ["0-3 мес", "3-6 мес", "6-12 мес", "1-3 года"])
                status = st.selectbox("Статус", ["идентифицирована", "квалифицирована", "предложение отправлено", "переговоры", "контракт", "исполнение", "завершён"])
                responsible = st.text_input("Ответственный от МИЭМ (ФИО)")
                docs_link = st.text_input("Ссылка на документы (URL)")
                submitted = st.form_submit_button("Сохранить гипотезу")
                if submitted and name and org_id and zadel_id:
                    session = Session()
                    z = session.query(Zadel).filter(Zadel.id == zadel_id).first()
                    new_h = Hypothesis(name=name, org_id=org_id, participation_type=part_type,
                                       miem_role_id=role_options.get(role_name), zadel_id=zadel_id, ugt=z.ugt,
                                       competitor=competitor, advantage=advantage, budget=budget if budget else None,
                                       horizon=horizon, status=status, responsible=responsible, docs_link=docs_link,
                                       created_at=str(date.today()))
                    session.add(new_h)
                    session.flush()
                    for bname in selected_barriers:
                        b = session.query(Barrier).filter(Barrier.name == bname).first()
                        if b:
                            session.add(HypothesisBarrier(hypothesis_id=new_h.id, barrier_id=b.id))
                    session.commit()
                    session.close()
                    st.success("Гипотеза создана!")
                    st.rerun()
                elif submitted and (not name or not org_id or not zadel_id):
                    st.error("Заполните обязательные поля")
        
        st.subheader("Список гипотез")
        session = Session()
        hypotheses = session.query(Hypothesis).all()
        if hypotheses:
            org_dict = get_org_dict()
            zadel_dict = get_zadel_dict()
            role_dict = get_role_dict()
            data = []
            for h in hypotheses:
                data.append({"ID": h.id, "Название": h.name, "Организация": org_dict.get(h.org_id, ""),
                             "Задел": zadel_dict.get(h.zadel_id, ""), "Роль МИЭМ": role_dict.get(h.miem_role_id, ""),
                             "Статус": h.status, "Бюджет": h.budget, "УГТ": h.ugt})
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
            
            selected_id = st.selectbox("Выберите ID гипотезы для просмотра/редактирования", [h.id for h in hypotheses])
            if selected_id:
                h = session.query(Hypothesis).filter(Hypothesis.id == selected_id).first()
                if h:
                    with st.expander("Редактировать гипотезу"):
                        with st.form("edit_hypothesis"):
                            new_name = st.text_input("Название", value=h.name)
                            new_part_type = st.radio("Тип участия", ["Заказчик", "Партнёр"], index=0 if h.participation_type == "Заказчик" else 1, horizontal=True)
                            
                            # Роль МИЭМ
                            role_options_list = list(role_dict.values())
                            current_role_name = role_dict.get(h.miem_role_id, role_options_list[0] if role_options_list else "")
                            try:
                                role_index = role_options_list.index(current_role_name)
                            except ValueError:
                                role_index = 0
                            new_role_name = st.selectbox("Роль МИЭМ", role_options_list, index=role_index)
                            
                            new_competitor = st.text_input("Конкурент", value=h.competitor or "")
                            new_advantage = st.text_area("Преимущество", value=h.advantage or "")
                            
                            session_b = Session()
                            barriers_all = session_b.query(Barrier).all()
                            barrier_names_all = [b.name for b in barriers_all]
                            current_barriers = session_b.query(HypothesisBarrier).filter(HypothesisBarrier.hypothesis_id == h.id).all()
                            current_barrier_names = []
                            for b in current_barriers:
                                b_obj = session_b.query(Barrier).get(b.barrier_id)
                                if b_obj:
                                    current_barrier_names.append(b_obj.name)
                            session_b.close()
                            new_barriers = st.multiselect("Барьеры", barrier_names_all, default=current_barrier_names)
                            
                            new_budget = st.number_input("Бюджет", value=h.budget or 0)
                            
                            # Горизонт
                            horizon_options = ["0-3 мес", "3-6 мес", "6-12 мес", "1-3 года"]
                            try:
                                horizon_index = horizon_options.index(h.horizon) if h.horizon else 0
                            except ValueError:
                                horizon_index = 0
                            new_horizon = st.selectbox("Горизонт", horizon_options, index=horizon_index)
                            
                            # Статус
                            status_options = ["идентифицирована", "квалифицирована", "предложение отправлено", "переговоры", "контракт", "исполнение", "завершён"]
                            try:
                                status_index = status_options.index(h.status) if h.status else 0
                            except ValueError:
                                status_index = 0
                            new_status = st.selectbox("Статус", status_options, index=status_index)
                            
                            new_responsible = st.text_input("Ответственный", value=h.responsible or "")
                            new_docs_link = st.text_input("Ссылка на документы", value=h.docs_link or "")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                save_h = st.form_submit_button("Сохранить изменения")
                            with col2:
                                delete_h = st.form_submit_button("Удалить гипотезу")
                            
                            if save_h:
                                h.name = new_name
                                h.participation_type = new_part_type
                                new_role_id = [k for k, v in role_dict.items() if v == new_role_name][0]
                                h.miem_role_id = new_role_id
                                h.competitor = new_competitor
                                h.advantage = new_advantage
                                h.budget = new_budget if new_budget else None
                                h.horizon = new_horizon
                                h.status = new_status
                                h.responsible = new_responsible
                                h.docs_link = new_docs_link
                                # Обновляем барьеры
                                session.query(HypothesisBarrier).filter(HypothesisBarrier.hypothesis_id == h.id).delete()
                                for bname in new_barriers:
                                    b = session.query(Barrier).filter(Barrier.name == bname).first()
                                    if b:
                                        session.add(HypothesisBarrier(hypothesis_id=h.id, barrier_id=b.id))
                                session.commit()
                                st.success("Гипотеза обновлена")
                                st.rerun()
                            if delete_h:
                                session.delete(h)
                                session.commit()
                                st.success("Гипотеза удалена")
                                st.rerun()
                    st.subheader(f"Детали гипотезы: {h.name}")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write(f"**Организация:** {org_dict.get(h.org_id)}")
                        st.write(f"**Задел:** {zadel_dict.get(h.zadel_id)} (УГТ {h.ugt})")
                        st.write(f"**Тип участия:** {h.participation_type}")
                        st.write(f"**Роль МИЭМ:** {role_dict.get(h.miem_role_id)}")
                        st.write(f"**Конкурент:** {h.competitor or '—'}")
                        st.write(f"**Преимущество:** {h.advantage or '—'}")
                    with col2:
                        st.write(f"**Статус:** {h.status}")
                        st.write(f"**Бюджет:** {h.budget or '—'}")
                        st.write(f"**Горизонт:** {h.horizon or '—'}")
                        st.write(f"**Ответственный:** {h.responsible or '—'}")
                        st.write(f"**Ссылка:** {h.docs_link or '—'}")
                    barriers_h = session.query(HypothesisBarrier).filter(HypothesisBarrier.hypothesis_id == h.id).all()
                    barrier_names_h = []
                    for bh in barriers_h:
                        b = session.query(Barrier).filter(Barrier.id == bh.barrier_id).first()
                        if b:
                            barrier_names_h.append(b.name)
                    st.write(f"**Барьеры:** {', '.join(barrier_names_h) if barrier_names_h else '—'}")
                    
                    st.subheader("Задачи")
                    tasks = session.query(Task).filter(Task.hypothesis_id == h.id).all()
                    if tasks:
                        task_data = []
                        for t in tasks:
                            task_data.append({"ID": t.id, "Задача": t.name, "Ответственный": t.responsible, "Срок": t.due_date, "Статус": t.status})
                        st.dataframe(pd.DataFrame(task_data), use_container_width=True)
                    else:
                        st.write("Нет задач.")
                    with st.expander("➕ Добавить задачу"):
                        with st.form("new_task"):
                            task_name = st.text_input("Название задачи*")
                            task_resp = st.text_input("Ответственный")
                            task_date = st.date_input("Срок выполнения", value=date.today())
                            submitted_task = st.form_submit_button("Сохранить задачу")
                            if submitted_task and task_name:
                                new_task = Task(hypothesis_id=h.id, name=task_name, responsible=task_resp, due_date=str(task_date), status="не выполнена")
                                session.add(new_task)
                                session.commit()
                                st.success("Задача добавлена")
                                st.rerun()
        else:
            st.info("Нет гипотез. Создайте первую через форму выше.")
        session.close()
    
    with tab_tasks:
        st.subheader("Задачи по всем гипотезам (не выполненные)")
        session = Session()
        tasks = session.query(Task).filter(Task.status != "выполнена").all()
        if tasks:
            hyp_dict = {h.id: h.name for h in session.query(Hypothesis).all()}
            for t in tasks:
                with st.expander(f"{t.name} (гипотеза: {hyp_dict.get(t.hypothesis_id, '?')})"):
                    st.write(f"**Ответственный:** {t.responsible or '—'}")
                    st.write(f"**Срок:** {t.due_date or '—'}")
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"✅ Выполнить", key=f"done_{t.id}"):
                            t.status = "выполнена"
                            session.commit()
                            st.rerun()
                    with col2:
                        if st.button(f"🗑 Удалить", key=f"del_{t.id}"):
                            session.delete(t)
                            session.commit()
                            st.rerun()
        else:
            st.info("Нет активных задач.")
        session.close()

# -------------------- ИМПОРТ ИЗ EXCEL --------------------
elif page == "Импорт из Excel":
    st.header("Импорт из Excel (одноразовый)")
    st.markdown("Загрузите файл Excel, содержащий лист **«Аналитика»**, полученный из макроса сбора паспортов.")
    
    uploaded_file = st.file_uploader("Выберите файл Excel", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name="Аналитика")
            st.success(f"Файл загружен. Найдено {len(df)} строк.")
            st.dataframe(df.head(10), use_container_width=True)
            
            if st.button("Начать импорт"):
                session = Session()
                stats = {"orgs": 0, "zadels": 0, "hypotheses": 0, "errors": 0}
                org_cache = {}
                zadel_cache = {}
                
                for idx, row in df.iterrows():
                    try:
                        org_name = row.get("Название заказчика")
                        if pd.isna(org_name) or not str(org_name).strip():
                            continue
                        org_name = str(org_name).strip()
                        if org_name not in org_cache:
                            org = session.query(Organization).filter(Organization.name == org_name).first()
                            if not org:
                                org = Organization(name=org_name, notes=row.get("Решаемая проблема") if not pd.isna(row.get("Решаемая проблема")) else None)
                                session.add(org)
                                session.flush()
                                stats["orgs"] += 1
                            org_cache[org_name] = org.id
                        org_id = org_cache[org_name]
                        
                        collective_name = row.get("Подразделение")
                        if not pd.isna(collective_name) and collective_name:
                            collective_name = str(collective_name).strip()
                        else:
                            collective_name = None
                        
                        grnti = row.get("Шифр ГРНТИ")
                        competence = row.get("Ключевая компетенция")
                        if pd.isna(grnti):
                            grnti = ""
                        else:
                            grnti = str(grnti).strip()
                        if pd.isna(competence):
                            competence = ""
                        else:
                            competence = str(competence).strip()
                        zadel_key = f"{grnti}|{competence[:100]}"
                        if zadel_key not in zadel_cache:
                            existing_zadel = session.query(Zadel).filter(Zadel.grnti == grnti, Zadel.competencies == competence).first()
                            if not existing_zadel:
                                team = row.get("Команда") if not pd.isna(row.get("Команда")) else ""
                                publications = row.get("Публикации") if not pd.isna(row.get("Публикации")) else ""
                                grants = row.get("Гранты") if not pd.isna(row.get("Гранты")) else ""
                                niokr = row.get("НИОКР") if not pd.isna(row.get("НИОКР")) else ""
                                rid_protected = row.get("Охраняемые РИД") if not pd.isna(row.get("Охраняемые РИД")) else ""
                                rid_unprotected = row.get("Неоформленные РИД") if not pd.isna(row.get("Неоформленные РИД")) else ""
                                ugt = row.get("УГТ")
                                if pd.isna(ugt):
                                    ugt = 1
                                else:
                                    ugt_str = str(ugt)
                                    match = re.search(r'\d+', ugt_str)
                                    ugt = int(match.group()) if match else 1
                                next_ugt_needs = row.get("Что нужно для следующего УГТ") if not pd.isna(row.get("Что нужно для следующего УГТ")) else ""
                                new_zadel = Zadel(
                                    name=f"{grnti} {competence[:50]}" if grnti else competence[:100],
                                    grnti=grnti, collective=collective_name, team=str(team), competencies=competence,
                                    publications=str(publications), grants=str(grants), niokr=str(niokr),
                                    rid_protected=str(rid_protected), rid_unprotected=str(rid_unprotected) if rid_unprotected else None,
                                    ugt=ugt, next_ugt_needs=str(next_ugt_needs)
                                )
                                session.add(new_zadel)
                                session.flush()
                                stats["zadels"] += 1
                                zadel_cache[zadel_key] = new_zadel.id
                            else:
                                zadel_cache[zadel_key] = existing_zadel.id
                        zadel_id = zadel_cache[zadel_key]
                        
                        hypothesis_name = f"{org_name} – {zadel_id}"
                        participation_type = "Заказчик"
                        miem_role = row.get("Роль МИЭМ")
                        if pd.isna(miem_role):
                            miem_role_id = None
                        else:
                            miem_role = str(miem_role).strip()
                            role_obj = session.query(MiemRole).filter(MiemRole.name == miem_role).first()
                            miem_role_id = role_obj.id if role_obj else None
                        competitor = row.get("Конкурент") if not pd.isna(row.get("Конкурент")) else ""
                        advantage = row.get("Преимущество") if not pd.isna(row.get("Преимущество")) else ""
                        horizon = row.get("Горизонт реализации") if not pd.isna(row.get("Горизонт реализации")) else ""
                        status = "идентифицирована"
                        responsible = row.get("Руководитель") if not pd.isna(row.get("Руководитель")) else ""
                        budget = None
                        docs_link = None
                        new_hyp = Hypothesis(
                            name=hypothesis_name, org_id=org_id, participation_type=participation_type,
                            miem_role_id=miem_role_id, zadel_id=zadel_id, ugt=ugt,
                            competitor=str(competitor), advantage=str(advantage), budget=budget,
                            horizon=str(horizon), status=status, responsible=str(responsible),
                            docs_link=docs_link, created_at=str(date.today())
                        )
                        session.add(new_hyp)
                        session.flush()
                        stats["hypotheses"] += 1
                        
                        barriers_str = row.get("Барьеры")
                        if not pd.isna(barriers_str) and barriers_str:
                            barrier_names = [b.strip() for b in str(barriers_str).split(',')]
                            for bname in barrier_names:
                                barrier = session.query(Barrier).filter(Barrier.name == bname).first()
                                if barrier:
                                    existing_link = session.query(HypothesisBarrier).filter(
                                        HypothesisBarrier.hypothesis_id == new_hyp.id,
                                        HypothesisBarrier.barrier_id == barrier.id
                                    ).first()
                                    if not existing_link:
                                        session.add(HypothesisBarrier(hypothesis_id=new_hyp.id, barrier_id=barrier.id))
                    except Exception as e:
                        stats["errors"] += 1
                        st.error(f"Ошибка в строке {idx+2}: {str(e)}")
                        continue
                session.commit()
                session.close()
                st.success(f"Импорт завершён. Создано: организаций – {stats['orgs']}, заделов – {stats['zadels']}, гипотез – {stats['hypotheses']}. Ошибок: {stats['errors']}.")
        except Exception as e:
            st.error(f"Ошибка при чтении файла: {str(e)}")