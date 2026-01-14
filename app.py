import time
import requests
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from urllib.parse import urlparse

# -----------------------
# Конфигурация API SpeedyIndex
# -----------------------
SPEEDY_BASE_URL = "https://api.speedyindex.com/v2"

# -----------------------
# Вспомогательные функции (Helpers)
# -----------------------
def get_headers(api_key):
    return {
        "Authorization": api_key,
        "Content-Type": "application/json"
    }

def get_balance(api_key):
    """Получаем баланс аккаунта"""
    try:
        url = f"{SPEEDY_BASE_URL}/account"
        resp = requests.get(url, headers=get_headers(api_key), timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            # SpeedyIndex возвращает баланс для indexer и checker отдельно
            # Нам нужен checker
            checker_bal = data.get("balance", {}).get("checker", 0)
            return checker_bal
    except Exception:
        return None
    return None

def send_slack_notification(token, channel, message):
    """Отправка уведомления в Slack"""
    url = "https://slack.com/api/chat.postMessage"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    payload = {
        "channel": channel,
        "text": message
    }
    try:
        requests.post(url, headers=headers, json=payload, timeout=10)
    except Exception as e:
        print(f"Slack error: {e}")

def find_header_row(ws, max_scan=20):
    """
    Ищем строку заголовков, где:
    - Колонка B (2) содержит 'Referring Page URL'
    Если не нашли, возвращаем 1
    """
    for r in range(1, min(ws.max_row, max_scan) + 1):
        val = ws.cell(row=r, column=2).value
        if isinstance(val, str) and "referring page url" in val.lower():
            return r
    return 1

def looks_like_url(val):
    if not isinstance(val, str):
        return False
    s = val.strip().lower()
    return s.startswith("http://") or s.startswith("https://")

# -----------------------
# Основной UI Streamlit
# -----------------------
st.set_page_config(page_title="SpeedyIndex Checker", layout="wide")
st.title("Проверка индексации (SpeedyIndex)")

# 1. Проверка Secrets
if "speedyindex" not in st.secrets or "slack" not in st.secrets:
    st.error("Не настроены secrets! Добавьте секции [speedyindex] и [slack].")
    st.stop()

api_key = st.secrets["speedyindex"]["api_key"]
slack_token = st.secrets["slack"]["bot_token"]
slack_channel = st.secrets["slack"]["channel_id"]

# 2. Отображение баланса
balance = get_balance(api_key)
if balance is not None:
    st.success(f"💰 Баланс SpeedyIndex (Checker): **{balance}** проверок")
else:
    st.warning("Не удалось получить баланс. Проверьте API ключ.")

st.markdown("---")

# 3. Загрузка файла
uploaded_file = st.file_uploader("Загрузите файл .xlsx", type=["xlsx"])

if uploaded_file:
    # Читаем файл в память
    wb = load_workbook(BytesIO(uploaded_file.getvalue()))
    all_sheet_names = wb.sheetnames
    
    selected_sheets = []

    # ЛОГИКА ВЫБОРА ЛИСТОВ
    if len(all_sheet_names) > 1:
        st.info(f"В файле найдено {len(all_sheet_names)} листов.")
        selected_sheets = st.multiselect(
            "Выберите листы для обработки:", 
            options=all_sheet_names,
            default=all_sheet_names
        )
    else:
        # Если лист один - выбираем его автоматически без вопросов
        selected_sheets = all_sheet_names

    if not selected_sheets:
        st.warning("Выберите хотя бы один лист для продолжения.")
        st.stop()

    if st.button("🚀 Начать проверку"):
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_sheets = len(selected_sheets)
        sheets_processed = 0
        total_links_checked = 0
        
        # Для уведомления в слак
        slack_report = []

        # Создаем сессию requests для переиспользования соединений
        session = requests.Session()
        session.headers.update(get_headers(api_key))

        for sheet_name in selected_sheets:
            status_text.write(f"⏳ Обработка листа: **{sheet_name}**...")
            ws = wb[sheet_name]
            
            # 1. Находим заголовки и данные
            header_row = find_header_row(ws)
            # Принудительно ставим заголовки для ясности результата
            ws.cell(row=header_row, column=4).value = "Index" # Column D
            
            urls_map = {} # { normalized_url : [list of row_indices] }
            raw_urls = [] # list for API
            
            # Собираем URL
            start_row = header_row + 1
            for r in range(start_row, ws.max_row + 1):
                cell_val = ws.cell(row=r, column=2).value
                if looks_like_url(cell_val):
                    clean_url = cell_val.strip()
                    raw_urls.append(clean_url)
                    
                    if clean_url not in urls_map:
                        urls_map[clean_url] = []
                    urls_map[clean_url].append(r)

            if not raw_urls:
                status_text.write(f"⚠️ На листе {sheet_name} не найдено ссылок.")
                continue

            # 2. Создаем задачу в SpeedyIndex
            # API принимает до 10k ссылок, мы отправляем весь лист сразу
            create_payload = {
                "title": f"Streamlit check {sheet_name}",
                "urls": raw_urls
            }
            
            try:
                # POST create task
                r_create = session.post(
                    f"{SPEEDY_BASE_URL}/task/google/checker/create", 
                    json=create_payload
                )
                res_create = r_create.json()
                
                if res_create.get("code") != 0:
                    st.error(f"Ошибка создания задачи для листа {sheet_name}: {res_create}")
                    continue
                
                task_id = res_create.get("task_id")
                status_text.write(f"Task ID: {task_id}. Ожидание результатов...")

                # 3. Полллинг статуса (ждем пока is_completed = true)
                is_completed = False
                attempts = 0
                while not is_completed and attempts < 60: # макс 3-4 минуты ожидания
                    time.sleep(3) # ждем 3 сек
                    
                    r_status = session.post(
                        f"{SPEEDY_BASE_URL}/task/google/checker/status", 
                        json={"task_ids": [task_id]}
                    )
                    res_status = r_status.json()
                    
                    task_info = res_status.get("result", [])[0]
                    if task_info.get("is_completed"):
                        is_completed = True
                    else:
                        attempts += 1
                        status_text.write(f"Лист {sheet_name}: Обработано {task_info.get('processed_count', 0)} из {task_info.get('size', 0)}...")

                if not is_completed:
                    st.error(f"Таймаут проверки листа {sheet_name}")
                    continue

                # 4. Получаем отчет (Report)
                r_report = session.post(
                    f"{SPEEDY_BASE_URL}/task/google/checker/report", 
                    json={"task_id": task_id}
                )
                data_report = r_report.json()
                
                # Списки ссылок из ответа
                indexed_list = set(data_report.get("result", {}).get("indexed_links", []))
                # Unindexed нам не обязателен для проверки "in", но он есть в data_report
                
                # 5. Записываем результаты в Excel
                # Проходим по всем URL, которые мы отправляли
                for url, rows in urls_map.items():
                    # Проверяем, есть ли url в списке проиндексированных
                    # SpeedyIndex может немного нормализовать ссылки, но обычно возвращает как есть
                    is_indexed = url in indexed_list
                    
                    for r_idx in rows:
                        ws.cell(row=r_idx, column=4).value = is_indexed
                
                count_indexed = len(indexed_list)
                count_total = len(raw_urls)
                slack_report.append(f"• List *{sheet_name}*: {count_indexed}/{count_total} indexed")
                total_links_checked += count_total
                
            except Exception as e:
                st.error(f"Критическая ошибка на листе {sheet_name}: {e}")

            sheets_processed += 1
            progress_bar.progress(sheets_processed / total_sheets)

        # -----------------------
        # Финализация
        # -----------------------
        status_text.success("✅ Проверка завершена!")
        
        # Сохранение файла
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Скачать результат (.xlsx)",
            data=output,
            file_name="checked_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Отправка в Slack
        if slack_report:
            msg_header = f"🤖 *Indexation Check Complete*\nTotal checked: {total_links_checked}\n\nDetails:\n"
            full_msg = msg_header + "\n".join(slack_report)
            send_slack_notification(slack_token, slack_channel, full_msg)
            st.toast("Уведомление отправлено в Slack", icon="📨")
