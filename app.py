import time
import requests
import streamlit as st
import pandas as pd
from io import BytesIO

# -----------------------
# Конфигурация
# -----------------------
SPEEDY_BASE_URL = "https://api.speedyindex.com/v2"

# -----------------------
# Функции
# -----------------------
def get_headers(api_key):
    return {
        "Authorization": api_key,
        "Content-Type": "application/json"
    }

def get_balance(api_key):
    try:
        url = f"{SPEEDY_BASE_URL}/account"
        resp = requests.get(url, headers=get_headers(api_key), timeout=5)
        if resp.status_code == 200:
            return resp.json().get("balance", {}).get("checker", 0)
    except:
        return None
    return None

def send_slack_notification(token, channel, message):
    try:
        requests.post(
            "https://slack.com/api/chat.postMessage",
            headers={"Authorization": f"Bearer {token}"},
            json={"channel": channel, "text": message},
            timeout=3
        )
    except:
        pass

def find_header_row_and_df(excel_file, sheet_name):
    """
    Быстро читает первые строки, чтобы найти, где начинаются заголовки (ищем 'Source', 'Link' и т.д.)
    Возвращает подготовленный DataFrame.
    """
    # Читаем первые 10 строк без заголовков
    preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=10)
    
    header_row_idx = 0
    found = False
    
    # Ищем строку, содержащую ключевые слова
    keywords = ['source', 'url', 'link', 'referring page']
    
    for idx, row in preview.iterrows():
        # Преобразуем строку в нижний регистр и ищем совпадения
        row_str = row.astype(str).str.lower().tolist()
        if any(k in ' '.join(row_str) for k in keywords):
            header_row_idx = idx
            found = True
            break
            
    # Если не нашли, пробуем 0-ю строку по умолчанию
    if not found:
        header_row_idx = 0

    # Читаем лист полностью уже с правильным заголовком
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row_idx)
    return df, header_row_idx

def looks_like_url(val):
    if not isinstance(val, str): return False
    return val.strip().lower().startswith(('http://', 'https://'))

# -----------------------
# UI Streamlit
# -----------------------
st.set_page_config(page_title="SpeedyIndex TURBO", layout="wide")
st.title("⚡ Проверка индексации (TURBO Mode)")

if "speedyindex" not in st.secrets or "slack" not in st.secrets:
    st.error("Нет секретов [speedyindex] или [slack]!")
    st.stop()

api_key = st.secrets["speedyindex"]["api_key"]
slack_token = st.secrets["slack"]["bot_token"]
slack_channel = st.secrets["slack"]["channel_id"]

# Баланс
bal = get_balance(api_key)
if bal is not None:
    st.success(f"💰 Баланс: {bal}")

uploaded_file = st.file_uploader("Файл .xlsx (Загрузка будет мгновенной)", type=["xlsx"])

if uploaded_file:
    # 1. Мгновенное чтение структуры через Pandas
    try:
        xl_file = pd.ExcelFile(uploaded_file)
        all_sheets = xl_file.sheet_names
    except Exception as e:
        st.error(f"Ошибка чтения файла: {e}")
        st.stop()

    # Выбор листов
    if len(all_sheets) > 1:
        selected_sheets = st.multiselect("Выберите листы:", all_sheets, default=all_sheets)
    else:
        selected_sheets = all_sheets

    if not selected_sheets:
        st.stop()

    if st.button("🚀 ЗАПУСК (TURBO)"):
        
        progress_bar = st.progress(0)
        status_box = st.empty()
        
        session = requests.Session()
        session.headers.update(get_headers(api_key))
        
        # Словарь для хранения результатов: {sheet_name: modified_dataframe}
        processed_sheets = {}
        
        # Активные задачи API
        active_tasks = {} # task_id -> {sheet_name, urls_list}
        total_urls_sent = 0
        
        # --- ЭТАП 1: Подготовка данных и отправка в API ---
        status_box.info("Чтение данных и отправка задач...")
        
        for sheet in selected_sheets:
            # Умный поиск заголовка и чтение данных
            df, _ = find_header_row_and_df(xl_file, sheet)
            
            # Ищем колонку Source (независимо от регистра)
            col_map = {c.lower(): c for c in df.columns}
            target_col = None
            for k in ['source', 'url', 'link', 'referring page url']:
                if k in col_map:
                    target_col = col_map[k]
                    break
            
            if not target_col:
                st.warning(f"На листе '{sheet}' не найдена колонка Source/URL. Пропускаем.")
                processed_sheets[sheet] = df # Сохраняем как есть
                continue

            # Фильтруем валидные URL для отправки
            # Создаем маску, чтобы потом записать ответы на свои места
            valid_mask = df[target_col].apply(looks_like_url)
            urls_to_check = df[target_col][valid_mask].tolist()
            urls_to_check = [u.strip() for u in urls_to_check]
            
            if not urls_to_check:
                processed_sheets[sheet] = df
                continue
                
            total_urls_sent += len(urls_to_check)
            
            # Отправка в API
            try:
                resp = session.post(
                    f"{SPEEDY_BASE_URL}/task/google/checker/create",
                    json={"title": sheet, "urls": urls_to_check},
                    timeout=10
                )
                data = resp.json()
                if data.get("code") == 0:
                    task_id = data["task_id"]
                    active_tasks[task_id] = {
                        "sheet": sheet,
                        "urls": urls_to_check, # Для контроля порядка (хотя API вернет список)
                        "original_df": df,
                        "valid_mask": valid_mask
                    }
                else:
                    st.error(f"Ошибка API (Лист {sheet}): {data}")
                    processed_sheets[sheet] = df 
            except Exception as e:
                st.error(f"Сбой сети (Лист {sheet}): {e}")
                processed_sheets[sheet] = df

        if not active_tasks:
            st.warning("Нет активных задач.")
            st.stop()

        # --- ЭТАП 2: Параллельное ожидание (Batch Wait) ---
        completed_ids = set()
        all_ids = list(active_tasks.keys())
        start_time = time.time()
        
        while len(completed_ids) < len(all_ids):
            if time.time() - start_time > 300: # 5 минут таймаут
                st.error("Таймаут ожидания API.")
                break
            
            pending = [tid for tid in all_ids if tid not in completed_ids]
            
            try:
                # Проверяем статус пачкой
                r = session.post(
                    f"{SPEEDY_BASE_URL}/task/google/checker/status",
                    json={"task_ids": pending}, timeout=10
                )
                tasks_status = r.json().get("result", [])
                
                still_running = 0
                for t_stat in tasks_status:
                    tid = t_stat["id"]
                    
                    if t_stat.get("is_completed"):
                        if tid not in completed_ids:
                            # Задача готова — получаем отчет
                            r_rep = session.post(
                                f"{SPEEDY_BASE_URL}/task/google/checker/report",
                                json={"task_id": tid}, timeout=15
                            )
                            rep_data = r_rep.json()
                            indexed_set = set(rep_data.get("result", {}).get("indexed_links", []))
                            
                            # --- ОБРАБОТКА РЕЗУЛЬТАТА ---
                            task_ctx = active_tasks[tid]
                            df = task_ctx["original_df"]
                            mask = task_ctx["valid_mask"]
                            
                            # Логика простановки TRUE/FALSE
                            # Мы используем .apply к колонке URL, проверяя наличие в indexed_set
                            target_col_name = df.columns[df.columns.str.lower().isin(['source', 'url', 'link'])][0]
                            
                            # Создаем серию результатов только для валидных строк
                            results_series = df.loc[mask, target_col_name].apply(
                                lambda x: (x.strip() in indexed_set) if isinstance(x, str) else False
                            )
                            
                            # Записываем в колонку Index (создаем новую или перезаписываем)
                            df.loc[mask, "Index"] = results_series
                            # Для невалидных URL можно оставить пустоту или False
                            
                            processed_sheets[task_ctx["sheet"]] = df
                            completed_ids.add(tid)
                    else:
                        still_running += 1
                
                # Обновление UI
                done = len(completed_ids)
                total = len(all_ids)
                progress_bar.progress(done / total)
                status_box.info(f"Проверка в процессе... Готово: {done}/{total}. В работе: {still_running}")
                
                if still_running > 0:
                    time.sleep(2.5) # Пауза между опросами
                    
            except Exception as e:
                st.error(f"Ошибка опроса API: {e}")
                time.sleep(5)

        # --- ЭТАП 3: Сохранение и отчет ---
        progress_bar.progress(1.0)
        status_box.success("Готово! Формируем файл...")
        
        # Сохранение через Pandas (очень быстро)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Проходим по всем листам (в том порядке, как они были в исходнике)
            for sheet_name in all_sheets:
                if sheet_name in processed_sheets:
                    # Записываем обработанный DF
                    processed_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    # Если лист не выбирали, можно попробовать сохранить старый (но это сложно без openpyxl)
                    # В режиме Turbo мы сохраняем только выбранные или пустые
                    pass
                    
        output.seek(0)
        
        # Slack
        msg = f"🚀 *SpeedyIndex Turbo Report*\nTotal URLs checked: {total_urls_sent}\nSheets processed: {len(processed_sheets)}"
        send_slack_notification(slack_token, slack_channel, msg)
        
        st.download_button(
            label="📥 Скачать результат (Fast .xlsx)",
            data=output,
            file_name="checked_turbo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
