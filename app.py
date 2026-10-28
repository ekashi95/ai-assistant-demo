import streamlit as st
import json
from docx import Document
from docx.shared import Inches
import io
import sys
from openai import OpenAI
import pandas as pd

client = OpenAI(
    api_key=st.secrets["OPENAI_API_KEY"] if "OPENAI_API_KEY" in st.secrets else None,
    base_url="https://api.proxyapi.ru/v1"
)

llm = "openai/gpt-4o-mini"

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
if "data_structure" not in st.session_state:
    st.session_state.data_structure = "Файл не загружен"
if "df" not in st.session_state:
    st.session_state.df = None

# ИИ (ЗДЕСЬ ПОДКЛЮЧАЕТСЯ ВАШЕ API) ---
def call_llm_assistant(user_input, context="", data_structure="Данные еще не загружены"):
    """
    Вызов LLM через ProxyAPI с передачей структуры загруженного файла.
    """
    if client is None:
        return {
            "type": "text",
            "content": "❌ **Ошибка авторизации:** API-ключ `OPENAI_API_KEY` не найден в настройках (Secrets).",
            "code": ""
        }

    system_prompt = f"""
    Ты — ИИ-соавтор научных статей и эксперт по анализу данных (Data Scientist).
    Твоя задача — помогать исследователю анализировать данные и писать черновик научной статьи на русском языке.
    
    СТРУКТУРА ЗАГРУЖЕННОГО ФАЙЛА ПОЛЬЗОВАТЕЛЯ:
    {data_structure}
    
    При написании Python-кода используй строго те названия колонок и типы данных, которые указаны выше! 
    Считай, что файл уже находится в памяти в виде датафрейма с именем `df`. Твой код должен работать напрямую с объектом `df`.
    
    Ты должен проанализировать запрос пользователя и решить, требуется ли написание Python-кода для обработки данных:
    1. Если пользователю нужен расчет, очистка, графики или манипуляции с данными, верни "type": "code". 
    2. Если пользователь задает теоретический вопрос, просит объяснить формулу или методологию БЕЗ написания кода, верни "type": "text".
    
    Ты ОБЯЗАН отвечать строго в формате JSON со следующей структурой:
    {{
        "type": "text" или "code",
        "content": "Твой развернутый текстовый ответ, пояснения, формулы в формате LaTeX (используй $...$ для инлайна и $$...$$ для блоков)",
        "code": "Полный, рабочий и безопасный Python-код для обработки df. Заполняй это поле ТОЛЬКО если type равен 'code'. В коде используй print() для вывода результатов."
    }}

    Если пользователь просит построить график, обязательно добавляй в конец кода строку plt.savefig('chart.png', bbox_inches='tight') и закрывай график через plt.close().
    """
    
    full_user_content = f"Контекст исследования: {context}\n\nЗапрос пользователя: {user_input}"
    
    try:
        response = client.chat.completions.create(
            model="openai/gpt-4o", 
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_user_content}
            ],
            temperature=0.2
        )
        
        # Исправление для форматов ProxyAPI/старых библиотек
        try:
            content_text = response.choices[0].message.content
        except (AttributeError, TypeError):
            content_text = response.choices[0]['message']['content']
            
        result_json = json.loads(content_text)
        return result_json

    except Exception as e:
        return {
            "type": "text",
            "content": f"❌ Ошибка ProxyAPI: {str(e)}.",
            "code": ""
        }

# --- 3. ФУНКЦИЯ ГЕНЕРАЦИИ ЧЕРНОВИКА СТАТЬИ (.DOCX) ---
def generate_docx_with_llm(chat_history, user_goal, data_structure, execution_log):
    """
    Финальный сборщик черновика .docx. Отправляет историю расчетов в LLM,
    генерирует текст и вставляет графики, если они были построены.
    """
    if client is None:
        st.error("Не удалось сгенерировать статью: отсутствует API-ключ в настройках Secrets.")
        return None

    # 1. Собираем логи вычислений и проверяем наличие сохраненного графика
    captured_results = "\n".join([f"Результат запуска {i+1}:\n{log}" for i, log in enumerate(execution_log)])
    if not captured_results:
        captured_results = "Код на данном этапе не выполнялся, только теоретическое обсуждение."
    
    chart_exists = os.path.exists("chart.png")

    # 2. Промпт для ИИ с требованием описать график
    academic_system_prompt = f"""
    Ты — выпускающий редактор научного журнала (Scopus Q1).
    Твоя задача — собрать связный, строгий черновик научной статьи на русском языке на основе материалов исследования.
    Используй академический стиль, избегай общих фраз. Опирайся ТОЛЬКО на реальные цели и результаты расчетов.
    
    График в исследовании: {"Был успешно построен и сохранен на диск под именем chart.png" if chart_exists else "Не строился"}.
    Если график строился, обязательно включи в раздел "results" его детальный академический разбор, описание трендов и зависимостей, которые видны из расчетов.
    
    Верни ответ строго в формате JSON со следующей структурой:
    {{
        "introduction": "Текст раздела 'Введение'. Опиши актуальность и цель работы на основе вводных данных пользователя.",
        "materials_and_methods": "Текст раздела 'Материалы и методы'. Опиши структуру выборки и математические/статистические методы, которые обсуждались.",
        "results": "Текст раздела 'Результаты и обсуждение'. Интегрируй РЕАЛЬНЫЕ цифры из логов вычислений и описание графика (если он есть)."
    }}
    """
    
    user_materials = f"ЦЕЛЬ: {user_goal}\nСТРУКТУРА ДАННЫХ: {data_structure}\nЛОГИ РАСЧЕТОВ:\n{captured_results}"
    
    try:
        response = client.chat.completions.create(
            model="openai/gpt-4o",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": academic_system_prompt},
                {"role": "user", "content": user_materials}
            ],
            temperature=0.3
        )
        
        # Исправление для шлюза ProxyAPI
        try:
            content_text = response.choices.message.content
        except (AttributeError, TypeError):
            content_text = response.choices['message']['content']
            
        article_json = json.loads(content_text)
    except Exception as e:
        article_json = {
            "introduction": f"Введение к исследованию: {user_goal}",
            "materials_and_methods": "Методы обработки данных включают программный анализ на языке Python.",
            "results": f"Произошла ошибка генерации текста через API: {str(e)}."
        }

    # 3. Записываем текст и графики в документ Word
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    
    doc.add_heading('Материалы к научной статье', level=1)
    doc.add_paragraph("Документ подготовлен интерактивным ИИ-ассистентом на основе проведенного анализа данных.")
    
    # Раздел 1. Введение
    doc.add_heading('1. Введение', level=2)
    doc.add_paragraph(article_json.get("introduction", ""))
    
    # Раздел 2. Методы
    doc.add_heading('2. Материалы и методы', level=2)
    doc.add_paragraph(article_json.get("materials_and_methods", ""))
    
    # Раздел 3. Результаты
    doc.add_heading('3. Результаты и обсуждение', level=2)
    doc.add_paragraph(article_json.get("results", ""))
    
    # ВСТАВКА ГРАФИКА: Если файл с графиком существует, добавляем его в Word
    if chart_exists:
        doc.add_paragraph() # Небольшой отступ
        # Добавляем картинку, ограничивая ширину (например, 5.5 дюймов для стандартного листа A4)
        doc.add_picture("chart.png", width=Inches(5.5))
        # Добавляем подпись к рисунку по ГОСТу
        caption = doc.add_paragraph("Рисунок 1 – График зависимости показателей, построенный в ходе автоматизированного анализа данных.")
        caption.style = 'Quote' # Оформляем как цитату/выделенный текст
        doc.add_paragraph()
    
    # Приложение с логами расчетов
    doc.add_heading('Приложение: Протокол вычислений (Сырые логи)', level=2)
    doc.add_paragraph(captured_results, style='Quote')

    # Сохраняем в память для скачивания
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# --- 4. ИНТЕРФЕЙС: ЛЕВАЯ ПАНЕЛЬ (БОКОВАЯ) ---
with st.sidebar:
    st.header("📂 1. Параметры исследования")
    uploaded_file = st.file_uploader("Загрузите файл с данными (CSV, Excel)", type=["csv", "xlsx"])
    user_goal = st.text_area("Описание данных и цель вашей работы:")
    
    if st.button("🚀 Инициализировать ассистента", use_container_width=True):
        if uploaded_file and user_goal:
            # Читаем файл в зависимости от формата
            try:
                if uploaded_file.name.endswith('.csv'):
                    st.session_state.df = pd.read_csv(uploaded_file)
                else:
                    st.session_state.df = pd.read_excel(uploaded_file)
                
                # Создаем текстовое описание структуры для ИИ
                buffer = io.StringIO()
                buffer.write("--- ТИПЫ КОЛОНОК ---\n")
                st.session_state.df.info(buf=buffer)
                buffer.write("\n--- ПЕРВЫЕ 3 СТРОКИ ДАННЫХ ---\n")
                buffer.write(st.session_state.df.head(3).to_string())
                
                st.session_state.data_structure = buffer.getvalue()
                
                # Сброс старых логов
                st.session_state.chat_history = []
                st.session_state.execution_log = []
                st.session_state.pending_code = None
                st.session_state.initialized = True
                
                # Первый запрос к ИИ с передачей структуры данных
                response = call_llm_assistant(
                    "Начало работы. Поприветствуй пользователя и предложи первый шаг анализа.", 
                    context=user_goal, 
                    data_structure=st.session_state.data_structure
                )
                
                if response["type"] == "code":
                    st.session_state.pending_code = response["code"]
                st.session_state.chat_history.append({
                    "role": "assistant", 
                    "content": response["content"], 
                    "code": response.get("code")
                })
                st.rerun()
                
            except Exception as e:
                st.error(f"Не удалось прочитать файл: {str(e)}")
        else:
            st.error("Пожалуйста, загрузите файл и заполните цель исследования.")

    # Кнопка скачивания статьи (остается без изменений)
    if st.session_state.initialized:
        st.write("---")
        st.header("📄 4. Экспорт результатов")
        if st.button("📦 Сформировать черновик статьи", use_container_width=True):
            if os.path.exists("chart.png"):
                os.remove("chart.png") # Удаляем старый график, если он остался на диске
            with st.spinner("ИИ анализирует ход исследования и пишет текст статьи..."):
                docx_data = generate_docx_with_llm(
                    st.session_state.chat_history, 
                    user_goal, 
                    st.session_state.data_structure,
                    st.session_state.execution_log
                )
                
                if docx_data:
                    st.download_button(
                        label="📥 Скачать черновик .docx",
                        data=docx_data,
                        file_name="scientific_draft.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )


# --- 5. ИНТЕРФЕЙС: ОСНОВНАЯ ЗОНА (ОБНОВЛЕНИЕ ДЛЯ ВЫЗОВА API) ---
if st.session_state.initialized:
    col_chat, col_control = st.columns([1.2, 1.0], gap="large")
    
    # --- ЛЕВАЯ КОЛОНКА: ДИАЛОГ С ИИ ---
    with col_chat:
        st.subheader("💬 Научная дискуссия")
        
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
                if "code" in msg and msg["code"]:
                    with st.expander("Посмотреть предложенный код"):
                        st.code(msg["code"], language="python")
        
        if not st.session_state.pending_code:
            user_message = st.chat_input("Задайте вопрос или дайте новую команду...")
            if user_message:
                st.session_state.chat_history.append({"role": "user", "content": user_message})
                st.rerun()
                
        if st.session_state.chat_history and st.session_state.chat_history[-1]["role"] == "user":
            last_query = st.session_state.chat_history[-1]["content"]
            with st.spinner("ИИ анализирует запрос..."):
                # Передаем и контекст, и структуру данных во все последующие вызовы
                response = call_llm_assistant(
                    last_query, 
                    context=user_goal, 
                    data_structure=st.session_state.data_structure
                )
                
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

    # --- ПРАВАЯ КОЛОНКА: БЕЗОПАСНОЕ ВЫПОЛНЕНИЕ НАД ТЕКУЩИМ DF ---
    with col_control:
        st.subheader("⚙️ Контур управления кодом")
        
        if st.session_state.pending_code:
            st.warning("⚠️ ИИ сгенерировал код для обработки ваших данных и ждет проверки.")
            st.code(st.session_state.pending_code, language="python")
            
            if st.button("🚀 Делай, что задумал (Запустить код)", use_container_width=True, type="primary"):
                with st.spinner("Выполнение расчетов и проверка модулей..."):
                    import sys
                    
                    old_stdout = sys.stdout
                    redirected_output = sys.stdout = io.StringIO()
                    missing_module = None
                    
                    try:
                        # 1. Предварительная мягкая проверка импортов без падения сервера
                        lines = st.session_state.pending_code.split('\n')
                        for line in lines:
                            line = line.strip()
                            if line.startswith('import ') or line.startswith('from '):
                                parts = line.split()
                                if len(parts) > 1:
                                    module_name = parts[1].split('.')[0]
                                    try:
                                        __import__(module_name)
                                    except ModuleNotFoundError:
                                        missing_module = module_name
                                        break
                        
                        # 2. Если какого-то модуля не хватает, прерываем выполнение с понятным текстом
                        if missing_module:
                            stdout_result = (
                                f"⚠️ Не удалось запустить код. ИИ использовал библиотеку `{missing_module}`, "
                                f"которая не установлена на сервере.\n"
                                f"Пожалуйста, попросите ИИ в поле ниже: 'Перепиши код без использования {missing_module}'."
                            )
                        else:
                            # Если всё на месте — выполняем код над датафреймом
                            local_vars = {"df": st.session_state.df, "pd": pd}
                            exec(st.session_state.pending_code, {}, local_vars)
                            
                            if "df" in local_vars:
                                st.session_state.df = local_vars["df"]
                                
                            stdout_result = redirected_output.getvalue()
                            if not stdout_result:
                                stdout_result = "Код выполнен успешно, но ничего не вывел через print()."
                                
                    except Exception as e:
                        stdout_result = f"Ошибка выполнения скрипта: {str(e)}"
                    finally:
                        sys.stdout = old_stdout
                
                st.session_state.execution_log.append(stdout_result)
                st.session_state.chat_history.append({
                    "role": "assistant", 
                    "content": f"**Результат шага обработки:**\n```\n{stdout_result}\n```"
                })
                st.session_state.pending_code = None
                st.rerun()
                
                st.session_state.execution_log.append(stdout_result)
                st.session_state.chat_history.append({
                    "role": "assistant", 
                    "content": f"**Код успешно выполнен.** Результаты расчетов зафиксированы:\n```\n{stdout_result}\n```"
                })
                st.session_state.pending_code = None
                st.rerun()
            
            st.write("---")
            user_correction = st.text_input("Введите замечание текстом:", key="correction_input")
            if st.button("🔧 Отправить правку к коду", use_container_width=True) and user_correction:
                st.session_state.chat_history.append({"role": "user", "content": f"Правка к коду: {user_correction}"})
                st.session_state.pending_code = None
                st.rerun()
        else:
            st.info("ℹ️ Сейчас активен режим теоретической консультации.")
            if st.session_state.execution_log:
                with st.expander("📊 Посмотреть логи прошлых вычислений", expanded=True):
                    for idx, log in enumerate(st.session_state.execution_log):
                        st.caption(f"Запуск №{idx+1}")
                        st.code(log)
