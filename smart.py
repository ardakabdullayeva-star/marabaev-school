import subprocess
import sys
import os
import json

# ==========================================
# АВТОМАТИЧЕСКАЯ УСТАНОВКА БИБЛИОТЕК
# ==========================================
required_libraries = ["streamlit", "pycryptodome", "python-docx", "google-genai"]
for lib in required_libraries:
    try:
        if lib == "pycryptodome": import Crypto
        elif lib == "python-docx": import docx
        elif lib == "google-genai": import google.genai
        else: __import__(lib)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", lib])

import streamlit as st
import sqlite3
import hashlib
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from google import genai
from google.genai import types

# Настройка страницы под глобальный бренд школы
st.set_page_config(
    page_title="Marabaev's Digital School (MDS)", 
    layout="wide", 
    page_icon="🏫"
)

# ==========================================
# СЕКРЕТНЫЙ ВШИТЫЙ API-КЛЮЧ
# ==========================================
БЕЗОПАСНЫЙ_API_КЛЮЧ = os.environ.get("GEMINI_API_KEY", "") 

# ==========================================
# БАЗА ДАННЫХ ДЛЯ КТП (КАЛЕНДАРНО-ТЕМАТИЧЕСКОЕ ПЛАНИРОВАНИЕ)
# ==========================================
KTP_DATA = {
    "Grade 10 (Action for Kazakhstan)": {
        "Science & Scientific Phenomena": {
            "topic": "Virtual Reality and its Applications",
            "objectives": "10.4.2.1 - Understand specific information and detail in extended texts on a range of familiar general and curricular topics;\n10.5.2.1 - Use a growing range of vocabulary, which is appropriate to topic and context, and clear symbols."
        },
        "The Fast Fashion Culture": {
            "topic": "Global Trends vs. Local Traditional Values",
            "objectives": "10.5.5.1 - Develop with minimal support coherent arguments supported by reasons, examples and evidence in a romance of written genres;\n10.6.7.1 - Use a wide variety of relative, demonstrative, indefinite pronouns."
        },
        "School Life & Learning": {
            "topic": "Effective Learning Strategies and Memory",
            "objectives": "10.1.6.1 - Organize and activate talk with peers to solve problems collaboratively;\n10.3.7.1 - Use appropriate subject-specific vocabulary and syntax to talk about a range of general and curricular topics."
        }
    },
    "Grade 11 (Aspect for Kazakhstan)": {
        "Making Living Space Better": {
            "topic": "Smart Home Technologies and Architecture",
            "objectives": "11.4.3.1 - Skim and scan a range of fiction and non-fiction texts for give message;\n11.5.1.1 - Plan, write, edit and proofread work at text level with minimal support."
        },
        "The Media and Presentation": {
            "topic": "Ethics in Journalism and Fake News Detection",
            "objectives": "11.2.5.1 - Recognize the attitude or opinion of the speaker in unsupported extended talk;\n11.3.3.1 - Explain and justify own and others’ point of view on a range of global problems."
        }
    }
}

# БАЗА ДАННЫХ ПОЛЬЗОВАТЕЛЕЙ И АРХИВА MDS
DB_FILE = "mds_marabaev_7.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT)")
    c.execute("""CREATE TABLE IF NOT EXISTS plans (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, type TEXT,
                    school TEXT, teacher TEXT, grade TEXT, topic TEXT, objs TEXT, 
                    lang TEXT, content TEXT)""")
    conn.commit()
    conn.close()

def make_hash(password): return hashlib.sha256(str.encode(password)).hexdigest()
def register_user(u, p):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    try: c.execute("INSERT INTO users (username, password) VALUES (?, ?)", (u, make_hash(p))); conn.commit(); s = True
    except: s = False
    conn.close(); return s
def login_user(u, p):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT id, username FROM users WHERE username = ? AND password = ?", (u, make_hash(p))); user = c.fetchone()
    conn.close(); return user

init_db()

# ==========================================
# ИИ ДВИЖОК MDS
# ==========================================
def ask_ai_or_template(mode, data, lang):
    if БЕЗОПАСНЫЙ_API_КЛЮЧ:
        try:
            client = genai.Client(api_key=БЕЗОПАСНЫЙ_API_КЛЮЧ)
            if mode == "KSP":
                prompt = f"Create KSP lesson plan Order 130. Language: {lang}. Topic: {data['topic']}. Grade: {data['grade']}. Objectives: {data['objs']}. Tasks: {data['tasks']}. Generate clean JSON only with keys: open_teacher, open_student, open_assess, mid_teacher, mid_student, mid_assess, end_teacher, end_student, end_assess, ref_teacher, ref_student, ref_assess, criteria_table."
            else:
                prompt = f"Create SOR assessment task based on Kazakhstan curriculum. Language: {lang}. Grade: {data['grade']}. Topic: {data['topic']}. Objectives: {data['objs']}. Generate clean JSON only with keys: task_description, questions, descriptors, max_points."
                
            response = client.models.generate_content(
                model='gemini-2.5-flash', contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            return json.loads(response.text.strip())
        except:
            pass

    if mode == "KSP":
        return {
            "open_teacher": "Greeting. Action: Teacher introduces the lesson topic and presents learning objectives. Warming-up activity: Brainstorming ideas.",
            "open_student": "Students respond to greeting, write down the topic, and share their initial thoughts during brainstorming.",
            "open_assess": "Formative tracking: Learners are awarded for active verbal responses. (No formal marks)",
            "mid_teacher": f"Methodological delivery: Teacher explains new active vocabulary and sets textbook tasks: {data['tasks']}. Differentiated approach for SEN included.",
            "mid_student": "Active learning: Students complete vocabulary exercises, read the assigned text, and complete tasks in pairs.",
            "mid_assess": "Assessment technique: Analytical descriptors check matching accuracy and sentence building (1-7 points).",
            "end_teacher": "Consolidation stage: Teacher sums up the main points, highlights success, and gives explanation of the homework task.",
            "end_student": "Reflection & Writing: Students make notes of the homework and evaluate their peer’s efforts.",
            "end_assess": "Feedback technique: 'Two stars and a wish' protocol implemented. (3 points)",
            "ref_teacher": "Teacher distributes self-assessment worksheets to evaluate individual comfort levels.",
            "ref_student": "Students fill out the reflection maps, identifying fields of success and doubt.",
            "ref_assess": "Self-assessment sheet saved in portfolio.",
            "criteria_table": "• Applies thematic vocabulary correctly (5 points)\n• Expresses arguments with valid text evidence (5 points)"
        }
    else:
        return {
            "task_description": f"Summative Assessment for the unit. Task: Reading comprehension and Writing task on topic '{data['topic']}'.",
            "questions": "Task 1. Read the text and identify if statements are True, False or Not Given (4 points).\nTask 2. Write a short essay response (80-100 words) using topic vocabulary (3 points).",
            "descriptors": "• Identifies specific arguments in text — 4 points\n• Uses appropriate grammatical links — 1 point\n• Demonstrates range of lexical units — 2 points",
            "max_points": "7 points"
        }

# ==========================================
# ОФОРМЛЕНИЕ ДОКУМЕНТОВ WORD (СТИЛЬ MDS)
# ==========================================
def set_cell_margins(cell, top=140, bottom=140, left=150, right=150):
    tc = cell._tcPr or cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}'); node.set(qn('w:w'), str(val)); node.set(qn('w:type'), 'dxa'); tcMar.append(node)
    tcPr.append(tcMar)

def create_ksp_docx(data):
    doc = Document()
    p = doc.add_paragraph()
    r = p.add_run(f"APPROVED BY _____________\n{data['school']}\nShort-term Lesson Plan: Английский язык")
    r.font.name = 'Arial'; r.font.size = Pt(11); r.bold = True
    
    t1 = doc.add_table(rows=0, cols=2); t1.style = 'Table Grid'
    rows = [("Teacher’s name:", data['teacher']), ("Class / Grade:", data['grade']), ("Lesson Topic:", data['topic']), ("Learning Objectives:", data['objs'])]
    for k, v in rows:
        row = t1.add_row(); row.cells[0].text = k; row.cells[1].text = v
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        set_cell_margins(row.cells[0]); set_cell_margins(row.cells[1])
        
    doc.add_paragraph().add_run("\nLesson Flow & Stages\n").bold = True
    t2 = doc.add_table(rows=1, cols=4); t2.style = 'Table Grid'
    for i, h in enumerate(["Stage / Time", "Teacher Actions", "Student Actions", "Assessment"]):
        t2.rows[0].cells[i].text = h; t2.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
        set_cell_margins(t2.rows[0].cells[i])
        
    stages = [
        ("Opening (5 min)", data['o_t'], data['o_s'], data['o_a']),
        ("Middle (25 min)", data['m_t'], data['m_s'], data['m_a']),
        ("End (10 min)", data['e_t'], data['e_s'], data['e_a']),
        ("Reflection (5 min)", data['r_t'], data['r_s'], data['r_a'])
    ]
    for stg, t, s, a in stages:
        row = t2.add_row()
        for idx, text in enumerate([stg, t, s, a]): row.cells[idx].text = text; set_cell_margins(row.cells[idx])
        
    bio = BytesIO(); doc.save(bio); return bio.getvalue()

# ==========================================
# ИНТЕРФЕЙС ШКОЛЬНОЙ ЭКОСИСТЕМЫ
# ==========================================
st.markdown("<h2 style='text-align: center; color: #1E3A8A; font-family: Arial;'>🏫 Marabaev's Digital School (MDS)</h2>", unsafe_allow_html=True)
st.markdown("<h5 style='text-align: center; color: #4B5563; margin-bottom: 25px;'>Цифровая ИИ-экосистема планирования Школы-лицея №7 им. Н.А. Марабаева</h5>", unsafe_allow_html=True)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_id = None
    st.session_state.username = ""

# БОКОВАЯ ПАНЕЛЬ MDS
with st.sidebar:
    st.markdown("<div style='background-color: #E0F2FE; padding: 15px; border-radius: 10px; border-left: 5px solid #0284C7;'><b>MDS Staff ID Authorization</b></div>", unsafe_allow_html=True)
    if not st.session_state.logged_in:
        auth = st.radio("Действие:", ["Авторизация", "Регистрация учителей"])
        u = st.text_input("Школьный логин (Почта):")
        p = st.text_input("Пароль доступа:", type="password")
        if auth == "Авторизация":
            if st.button("🔐 Войти в MDS", use_container_width=True):
                user = login_user(u, p)
                if user:
                    st.session_state.logged_in, st.session_state.user_id, st.session_state.username = True, user[0], user[1]
                    st.rerun()
                else: st.error("Неверный ID сотрудника школы №7")
        else:
            if st.button("📝 Создать аккаунт", use_container_width=True):
                if u and p and register_user(u, p): st.success("Успешно добавлено в базу MDS!")
                else: st.error("Этот ID уже занят")
    else:
        st.success(f"👨‍🏫 Преподаватель: {st.session_state.username}")
        st.info("Организация: Школа-лицей №7")
        if st.button("🚪 Выйти из MDS", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns(3)
    with col2:
        st.info("💡 Добро пожаловать в единую систему генерации КСП и СОР/СОЧ Marabaev's Digital School. Авторизуйтесь на боковой панели.")
else:
    # РАБОЧИЕ СЕКЦИИ
    tab_ksp, tab_sor, tab_archive = st.tabs(["📋 Конструктор КСП (Приказ №130)", "🎯 Модуль СОР / СОЧ", "🗂 Внутришкольный Архив MDS"])
    
    # ----------------------------------------
    # ВКЛАДКА 1: КСП
    # ----------------------------------------
    with tab_ksp:
        st.markdown("<h3 style='color: #1E3A8A;'>Автозаполнение на базе школьного КТП</h3>", unsafe_allow_html=True)
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            sel_grade = st.selectbox("🎯 Выберите параллель классов:", list(KTP_DATA.keys()))
        with col_s2:
            sel_unit = st.selectbox("📖 Выберите сквозную тему / Раздел:", list(KTP_DATA[sel_grade].keys()))
            
        auto_topic = KTP_DATA[sel_grade][sel_unit]["topic"]
        auto_objs = KTP_DATA[sel_grade][sel_unit]["objectives"]
        
        st.success(f"**Выбрано из КТП школы:**\n\n*Тема урока:* {auto_topic}\n\n*Цели обучения:* \n{auto_objs}")
        
        st.markdown("---")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            school = st.text_input("Организация образования (Школа):", value="КГУ Школа-лицей №7 им. Н.А. Марабаева")
            teacher = st.text_input("ФИО преподавателя:", value=st.session_state.username)
            lang = st.selectbox("Язык документов:", ["English", "Русский", "Қазақша"], key="ksp_lang")
        with col_p2:
            tasks = st.text_area("📖 Упражнения из учебника:", value="Student's Book: Exercise 3, page 44. Analysis of infographic data.")
            idea = st.text_area("🚀 Пожелания / Методические идеи:", value="Сделать упор на парную интерактивную работу и активные методы.")
            
        if st.button("⚡ Сформировать официальный КСП", type="primary", use_container_width=True):
            inp = {"grade": sel_grade, "topic": auto_topic, "objs": auto_objs, "tasks": tasks, "idea": idea}
            with st.spinner("MDS ИИ выстраивает структуру урока по критериям..."):
                res = ask_ai_or_template("KSP", inp, lang)
                
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("INSERT INTO plans (user_id, type, school, teacher, grade, topic, objs, lang, content) VALUES (?, 'KSP', ?, ?, ?, ?, ?, ?, ?)",
                      (st.session_state.user_id, school, teacher, sel_grade, auto_topic, auto_objs, lang, json.dumps(res, ensure_ascii=False)))
            conn.commit(); conn.close()
            st.success("🎉 План успешно отправлен во внутришкольный архив MDS!")

    # ----------------------------------------
    # ВКЛАДКА 2: СОР
    # ----------------------------------------
    with tab_sor:
        st.markdown("<h3 style='color: #1E3A8A;'>Разработка суммативного оценивания (СОР) для MDS</h3>", unsafe_allow_html=True)
        
        col_sor1, col_sor2 = st.columns(2)
        with col_sor1:
            s_grade = st.selectbox("Параллель для СОР:", list(KTP_DATA.keys()), key="sor_g")
            s_unit = st.selectbox("Раздел для СОР:", list(KTP_DATA[s_grade].keys()), key="sor_u")
        with col_sor2:
            s_lang = st.selectbox("Язык СОР:", ["English", "Русский", "Қазақша"], key="sor_l")
            s_school = st.text_input("Школа для СОР:", value="КГУ Школа-лицей №7 им. Н.А. Марабаева", key="sor_sch")
            
        s_topic = KTP_DATA[s_grade][s_unit]["topic"]
        s_objs = KTP_DATA[s_grade][s_unit]["objectives"]
        
        if st.button("🛠 Разработать спецификацию и задания СОР", type="primary", use_container_width=True):
            inp = {"grade": s_grade, "topic": s_topic, "objs": s_objs}
            with st.spinner("MDS ИИ формирует критериальные задания..."):
                res = ask_ai_or_template("SOR", inp, s_lang)
                
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("INSERT INTO plans (user_id, type, school, teacher, grade, topic, objs, lang, content) VALUES (?, 'SOR', ?, ?, ?, ?, ?, ?, ?)",
                      (st.session_state.user_id, s_school, st.session_state.username, s_grade, s_topic, s_objs, s_lang, json.dumps(res, ensure_ascii=False)))
            conn.commit(); conn.close()
            st.success("🎯 Задания СОР успешно добавлены в ваш архив MDS!")

    # ----------------------------------------
    # ВКЛАДКА 3: АРХИВ MDS
    # ----------------------------------------
    with tab_archive:
        st.markdown("<h3 style='color: #1E3A8A;'>Цифровой архив документов MDS</h3>", unsafe_allow_html=True)
        
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT id, type, grade, topic, lang, content, school, teacher, objs FROM plans WHERE user_id = ?", (st.session_state.user_id,))
        records = c.fetchall(); conn.close()
        
        if not records:
            st.info("Архив MDS пока пуст. Сгенерируйте первый КСП или СОР.")
        else:
            for r_id, r_type, r_grade, r_topic, r_lang, r_json, r_school, r_teacher, r_objs in records:
                doc_data = json.loads(r_json)
                type_label = "📋 КСП" if r_type == "KSP" else "🎯 СОР"
                
                with st.expander(f"{type_label} — {r_grade} — {r_topic} [{r_lang}]"):
                    if r_type == "KSP":
                        st.write(f"**Организация:** {r_school} | **Учитель:** {r_teacher}")
                        st.table({
                            "Этап урока / Время": ["Открытие (5 мин)", "Основной (25 min)", "Закрепление (10 min)", "Рефлексия (5 min)"],
                            "Действия учителя": [doc_data.get('open_teacher',''), doc_data.get('mid_teacher',''), doc_data.get('end_teacher',''), doc_data.get('ref_teacher','')],
                            "Действия учащихся": [doc_data.get('open_student',''), doc_data.get('mid_student',''), doc_data.get('end_student',''), doc_data.get('ref_student','')],
                            "Критериальное оценивание": [doc_data.get('open_assess',''), doc_data.get('mid_assess',''), doc_data.get('end_assess',''), doc_data.get('ref_assess','')]
                        })
                        
                        w_data = {
                            'school': r_school, 'teacher': r_teacher, 'grade': r_grade, 'topic': r_topic, 'objs': r_objs,
                            'o_t': doc_data.get('open_teacher',''), 'o_s': doc_data.get('open_student',''), 'o_a': doc_data.get('open_assess',''),
                            'm_t': doc_data.get('mid_teacher',''), 'm_s': doc_data.get('mid_student',''), 'm_a': doc_data.get('mid_assess',''),
                            'e_t': doc_data.get('end_teacher',''), 'e_s': doc_data.get('end_student',''), 'e_a': doc_data.get('end_assess',''),
                            'r_t': doc_data.get('ref_teacher',''), 'r_s': doc_data.get('ref_student',''), 'r_a': doc_data.get('ref_assess','')
                        }
                        bytes_docx = create_ksp_docx(w_data)
                        st.download_button("📥 Скачать официальный файл КСП (Word .docx)", bytes_docx, file_name=f"KSP_{r_topic}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_ksp_{r_id}")
                    
                    else:
                        st.write(f"### Спецификация суммативного оценивания")
                        st.info(f"**Описание суммативной работы:**\n{doc_data.get('task_description','')}")
                        st.warning(f"**Текст тестовых заданий / Вопросы:**\n\n{doc_data.get('questions','')}")
                        st.success(f"**Рубрика выставления баллов и дескрипторы:**\n\n{doc_data.get('descriptors','')}\n\n**Максимальный балл:** {doc_data.get('max_points','')}")