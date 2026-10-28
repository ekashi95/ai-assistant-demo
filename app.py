import streamlit as st
from docx import Document
import io
import sys

# Настройка страницы Streamlit
st.set_page_config(layout="wide", page_title="Научный ИИ-Ассистент")
st.title("🔬 Научный ИИ-Ассистент")

# --- 1. ИНИЦИАЛИЗАЦИЯ СОСТОЯНИЯ СЕССИИ ---
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "pending_code" not in st.session_state:
    st.session_state.pending_code = None
if "execution_log" not in st.session_state:
    st.session_state.execution_log = []
if "initialized" not in st.session_state:
    st.session_state.initialized = False

# --- 2. СИМУЛЯЦИЯ ИИ (ЗДЕСЬ ПОДКЛЮЧАЕТСЯ ВАШЕ API) ---
def call_llm_assistant(user_input, context=""):
    """
    В реальном проекте здесь будет вызов OpenAI API или Gemini API.
    Модель настраивается системным промптом возвращать JSON со структурой:
    { "type": "text"|"code", "content": "текст", "code": "код если есть" }
    """
    ui_lower = user_input.lower()
    
    # Сценарий 1: Запрос на теорию, формулы или консультацию (Без кода)
    if any(word in ui_lower for word in ["формул", "посоветуй", "объясни", "какой метод"]):
        return {
            "type": "text",
            "content": (
                "Для проверки связи между переменными в вашем исследовании я рекомендую использовать "
                "**коэффициент ранговой корреляции Спирмена ($r_s$)**, так как медицинские данные часто "
                "имеют распределение, отличное от нормального. Формула расчета:\n\n"
                "$$r_s = 1 - \\frac{6 \\sum d_i^2}{n(n^2 - 1)}$$\n\n"
                "Где $d_i$ — разность между рангами показателей, а $n$ — объем выборки. "
                "Если вы согласны, мы можем написать скрипт для расчета этой матрицы."
            )
        }
    
    # Сценарий 2: Запрос на исправление кода или использование других параметров
    elif any(word in ui_lower for word in ["замени", "исправь", "ошибка", "правка к коду"]):
        return {
            "type": "code",
            "content": "Я скорректировал логику обработки данных согласно вашему замечанию. Теперь мы используем расчет медианы вместо среднего значения, чтобы исключить влияние выбросов.",
            "code": (
                "import pandas as pd\n"
                "import numpy as np\n\n"
                "# Расчет медианного значения и стандартного отклонения\n"
                "print('--- Результаты измененной обработки ---')\n"
                "print('Применена формула расчета медианных значений для групп.')\n"
                "print('Медиана группы А: 42.5 (ед.)')\n"
                "print('Медиана группы Б: 31.2 (ед.)')\n"
                "print('Различия статистически значимы, p < 0.05')"
            )
        }
        
    # Сценарий 3: Базовый шаг обработки данных
    else:
        return {
            "type": "code",
            "content": "Я проанализировал структуру вашего файла. Предлагаю начать с этапа предобработки: очистки данных от пустых значений и вывода базовой описательной статистики.",
            "code": (
                "import pandas as pd\n\n"
                "# Эмуляция загрузки и первичного анализа датасета\n"
                "print('--- Результаты первичного анализа ---')\n"
                "print('Успешно обработано строк: 1,240')\n"
                "print('Обнаружено и удалено пустых строк: 14')\n"
                "print('Базовые метрики распределения успешно рассчитаны.')"
            )
        }

# --- 3. ФУНКЦИЯ ГЕНЕРАЦИИ ЧЕРНОВИКА СТАТЬИ (.DOCX) ---
def generate_docx(chat_history, goal_text):
    doc = Document()
    
    # Задаем базовый шрифт для академического стиля
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    
    doc.add_heading('Черновик научной статьи', level=1)
    
    # Раздел 1. Введение
    doc.add_heading('1. Введение', level=2)
    doc.add_paragraph(
        f"Настоящая работа посвящена исследованию актуальной научной проблемы, описанной автором. "
        f"В рамках проведенного анализа ставилась следующая ключевая цель: {goal_text}. "
        f"Актуальность темы обусловлена необходимостью автоматизации и повышения точности интерпретации экспериментальных данных."
    )
    
    # Раздел 2. Литературный обзор
    doc.add_heading('2. Краткий обзор подходов', level=2)
    doc.add_paragraph(
        "В современной литературе для решения подобных задач применяются как классические статистические критерии, "
        "так и современные методы интеллектуального анализа данных. Рассматриваемый в работе подход "
        "нацелен на оптимизацию процесса обработки с сохранением строгого контроля со стороны исследователя."
    )
    
    # Раздел 3. Методология (Собираем все, что одобрил пользователь)
    doc.add_heading('3. Материалы и методы', level=2)
    doc.add_paragraph("Для достижения поставленной цели была реализована программная обработка данных на языке Python. Этапы исследования включали:")
    
    has_methods = False
    for msg in chat_history:
        if msg["role"] == "assistant" and "code" in msg:
            doc.add_paragraph(f"• {msg['content']}")
            doc.add_paragraph(f"Примененный программный алгоритм:\n{msg['code']}", style='Quote')
            has_methods = True
            
    if not has_methods:
        doc.add_paragraph("На данном этапе производилось теоретическое обсуждение методологии анализа без выполнения программного кода.")
        
    # Раздел 4. Результаты
    doc.add_heading('4. Результаты и обсуждение', level=2)
    doc.add_paragraph("На основе примененных алгоритмов были зафиксированы следующие промежуточные научные результаты:")
    
    if st.session_state.execution_log:
        for log in st.session_state.execution_log:
            doc.add_paragraph(log, style='Intense Quote')
    else:
        doc.add_paragraph("Экспериментальные расчеты находятся в процессе утверждения.")
        
    # Сохраняем в память
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# --- 4. ИНТЕРФЕЙС: ЛЕВАЯ ПАНЕЛЬ (БОКОВАЯ) ---
with st.sidebar:
    st.header("📂 1. Параметры исследования")
    uploaded_file = st.file_uploader("Загрузите файл с данными (CSV, Excel)", type=["csv", "xlsx"])
    user_goal = st.text_area("Описание данных и цель вашей работы:", placeholder="Например: Изучаю влияние фактора Х на показатель Y в группе пациентов...")
    
    if st.button("🚀 Инициализировать ассистента", use_container_width=True):
        if uploaded_file and user_goal:
            st.session_state.chat_history = []
            st.session_state.execution_log = []
            st.session_state.pending_code = None
            st.session_state.initialized = True
            
            # Первый шаг ИИ
            response = call_llm_assistant("Начало работы")
            if response["type"] == "code":
                st.session_state.pending_code = response["code"]
                st.session_state.chat_history.append({
                    "role": "assistant", 
                    "content": response["content"], 
                    "code": response["code"]
                })
            st.rerun()
        else:
            st.error("Пожалуйста, загрузите файл и заполните цель исследования.")

    # Кнопка скачивания черновика (доступна, если сессия начата)
    if st.session_state.initialized:
        st.write("---")
        st.header("📄 4. Экспорт результатов")
        if st.button("📦 Сформировать черновик статьи", use_container_width=True, type="secondary"):
            docx_data = generate_docx(st.session_state.chat_history, user_goal)
            st.download_button(
                label="📥 Скачать черновик .docx",
                data=docx_data,
                file_name="scientific_draft.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )


# --- 5. ИНТЕРФЕЙС: ОСНОВНАЯ ЗОНА (ДВЕ КОЛОНКИ) ---
if not st.session_state.initialized:
    st.info("👋 Добро пожаловать! Загрузите данные и укажите цель исследования на панели слева, чтобы начать работу с ИИ-ассистентом.")
else:
    col_chat, col_control = st.columns([1.2, 1.0], gap="large")
    
    # --- ЛЕВАЯ КОЛОНКА: ДИАЛОГ С ИИ ---
    with col_chat:
        st.subheader("💬 Научная дискуссия")
        
        # Отображение истории чата
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
                if "code" in msg:
                    with st.expander("Посмотреть предложенный код"):
                        st.code(msg["code"], language="python")
        
        # Поле ввода доступно ТОЛЬКО если код не ждет аппрува в правой колонке
        if not st.session_state.pending_code:
            user_message = st.chat_input("Задайте вопрос или дайте новую команду...")
            if user_message:
                st.session_state.chat_history.append({"role": "user", "content": user_message})
                st.rerun()
                
        # Если новое сообщение добавлено пользователем, но ИИ еще не ответил
        if st.session_state.chat_history and st.session_state.chat_history[-1]["role"] == "user":
            last_query = st.session_state.chat_history[-1]["content"]
            with st.spinner("ИИ анализирует запрос..."):
                response = call_llm_assistant(last_query)
                
            if response["type"] == "code":
                st.session_state.pending_code = response["code"]
                st.session_state.chat_history.append({
                    "role": "assistant", 
                    "content": response["content"], 
                    "code": response["code"]
                })
            else:
                st.session_state.chat_history.append({
                    "role": "assistant", 
                    "content": response["content"]
                })
            st.rerun()

# --- ПРАВАЯ КОЛОНКА: УПРАВЛЕНИЕ И ВЫПОЛНЕНИЕ ---
    with col_control:
        st.subheader("⚙️ Контур управления кодом")
        
        if st.session_state.pending_code:
            st.warning("⚠️ ИИ сгенерировал код для обработки ваших данных и ждет проверки.")
            st.code(st.session_state.pending_code, language="python")
            
            # Кнопка подтверждения запуска
            if st.button("🚀 Делай, что задумал (Запустить код)", use_container_width=True, type="primary"):
                with st.spinner("Выполнение расчетов на сервере..."):
                    # Перехватываем stdout (команды print)
                    old_stdout = sys.stdout
                    redirected_output = sys.stdout = io.StringIO()
                    
                    try:
                        # В продакшене здесь вызов безопасной песочницы (Docker / E2B / API)
                        exec(st.session_state.pending_code)
                        stdout_result = redirected_output.getvalue()
                    except Exception as e:
                        stdout_result = f"Ошибка выполнения скрипта: {str(e)}"
                    finally:
                        sys.stdout = old_stdout
                
                # Добавляем результаты выполнения в лог и чат
                st.session_state.execution_log.append(stdout_result)
                st.session_state.chat_history.append({
                    "role": "assistant", 
                    "content": f"**Код успешно выполнен.** Результаты расчетов зафиксированы:\n```\n{stdout_result}\n```"
                })
                st.session_state.pending_code = None # Снимаем блокировку чата
                st.rerun()
            
            # Поле для отправки текстовых правок к коду
            st.write("---")
            st.write("💡 **Не согласны с кодом?** Напишите замечание текстом ниже:")
            user_correction = st.text_input("Например: 'Используй тут медиану вместо среднего'", key="correction_input")
            
            if st.button("🔧 Отправить правку к коду", use_container_width=True) and user_correction:
                st.session_state.chat_history.append({"role": "user", "content": f"Правка к коду: {user_correction}"})
                st.session_state.pending_code = None # Сбрасываем старый код, ждем новый от ИИ
                st.rerun()
                
        else:
            st.info("ℹ️ Сейчас активен режим теоретической консультации. Вы можете свободно общаться в чате слева. Код для выполнения не сгенерирован.")
            
            # Вывод логов прошлых запусков, если они есть
            if st.session_state.execution_log:
                with st.expander("📊 Посмотреть логи прошлых вычислений", expanded=True):
                    for idx, log in enumerate(st.session_state.execution_log):
                        st.caption(f"Запуск №{idx+1}")
                        st.code(log)