import time
import re
from io import BytesIO
from urllib.parse import urlparse, urlunparse

import requests
import streamlit as st
from openpyxl import load_workbook

# -----------------------
# DataForSEO endpoints (V3)
# POST: без /advanced
# GET: с /advanced/{task_id}
# -----------------------
TASK_POST = "/v3/serp/google/organic/task_post"
TASKS_READY = "/v3/serp/google/organic/tasks_ready"
TASK_GET_ADV = "/v3/serp/google/organic/task_get/advanced/{task_id}"

# -----------------------
# Helpers
# -----------------------
def looks_like_url(v) -> bool:
    if not isinstance(v, str):
        return False
    s = v.strip().lower()
    return s.startswith("http://") or s.startswith("https://")

def norm_url(u: str) -> str:
    """Сравнение URL: без схемы, без www, без query, без trailing slash."""
    p = urlparse(u.strip())
    netloc = (p.netloc or "").lower()
    if netloc.startswith("www."):
        netloc = netloc[4:]
    path = (p.path or "").rstrip("/")
    return urlunparse(("", netloc, path, "", "", "")).lower()

def build_site_query(url: str) -> str:
    """
    Запрос для проверки индексации через SERP:
    https://site.com/path?x=1 -> site:site.com/path
    главная -> site:site.com
    """
    p = urlparse(url.strip())
    host = (p.netloc or "").lower()
    if host.startswith("www."):
        host = host[4:]
    path = (p.path or "").strip()
    path = "/" + path.lstrip("/")
    path = path.rstrip("/")
    if path in ("", "/"):
        return f"site:{host}"
    return f"site:{host}{path}"

def safe_json(resp: requests.Response):
    try:
        return resp.json()
    except Exception:
        raise RuntimeError(f"Ответ не JSON. HTTP={resp.status_code}. Body(first 500)={resp.text[:500]}")

def api_request(session: requests.Session, method: str, url: str, *, json_body=None, timeout=60, tries=6):
    delay = 1.0
    last_err = None
    for _ in range(tries):
        try:
            if method == "GET":
                r = session.get(url, timeout=timeout)
            else:
                r = session.post(url, json=json_body, timeout=timeout)

            # 429/5xx — подождём и повторим
            if r.status_code in (429, 500, 502, 503, 504):
                last_err = f"HTTP {r.status_code}"
                time.sleep(delay)
                delay = min(delay * 2, 20)
                continue

            return r
        except Exception as e:
            last_err = str(e)
            time.sleep(delay)
            delay = min(delay * 2, 20)

    raise RuntimeError(f"API request failed after retries: {last_err}")

def find_headers(ws, max_rows=10):
    """
    Ищем строку заголовков, где:
    - в колонке B стоит 'Referring Page URL'
    - в колонке D стоит 'Index'
    Если не нашли — считаем header_row=1 и принудительно ставим заголовки.
    """
    for r in range(1, min(ws.max_row or 1, max_rows) + 1):
        b = ws.cell(row=r, column=2).value
        d = ws.cell(row=r, column=4).value
        b_ok = isinstance(b, str) and b.strip().lower() == "referring page url"
        d_ok = isinstance(d, str) and d.strip().lower() == "index"
        if b_ok:
            return r  # достаточно найти B
    return 1

def collect_tasks_from_workbook(wb):
    """
    Возвращает:
    - tasks_meta: список (sheet_name, row, url)
    - total_urls
    И гарантирует, что в каждом листе стоят заголовки в B(header) и D(header).
    """
    tasks = []
    for ws in wb.worksheets:
        header_row = find_headers(ws)

        # гарантируем заголовки
        ws.cell(row=header_row, column=2).value = "Referring Page URL"
        ws.cell(row=header_row, column=4).value = "Index"

        # данные начинаются со следующей строки
        start_row = header_row + 1
        max_row = ws.max_row or start_row

        for r in range(start_row, max_row + 1):
            v = ws.cell(row=r, column=2).value  # колонка B
            if not looks_like_url(v):
                continue
            tasks.append((ws.title, r, v.strip()))

    return tasks, len(tasks)

def match_indexed(original_url: str, items):
    """indexed=True если нашли URL в organic результатах."""
    orig = norm_url(original_url)
    for it in items:
        if it.get("type") != "organic":
            continue
        u = it.get("url")
        if not u:
            continue
        if norm_url(u) == orig:
            return True
    return False

# -----------------------
# Streamlit UI
# -----------------------
st.set_page_config(page_title="Indexation Checker (DataForSEO)", layout="wide")
st.title("Проверка индексации ссылок (xlsx → xlsx)")

st.caption(
    "Берём ссылки из колонки B (Referring Page URL) на каждом листе и пишем результат в колонку D (Index)."
)

uploaded = st.file_uploader("Загрузи .xlsx", type=["xlsx"])

# Настройки запроса (можно оставить дефолт)
with st.expander("Настройки SERP (опционально)", expanded=False):
    col1, col2, col3 = st.columns(3)
    location_code = col1.number_input("location_code", value=2840, min_value=1)  # по умолчанию US
    language_code = col2.text_input("language_code", value="en")
    depth = col3.number_input("depth", value=10, min_value=1, max_value=100)
    device = st.selectbox("device", ["desktop", "mobile"], index=0)

# Secrets
if "dataforseo" not in st.secrets:
    st.error("Не найдены secrets: [dataforseo]. Добавь их в Streamlit Cloud → Settings → Secrets.")
    st.stop()

dfseo = st.secrets["dataforseo"]
login = dfseo.get("login")
password = dfseo.get("password")
host = (dfseo.get("host") or "api.dataforseo.com").replace("https://", "").replace("http://", "").strip().strip("/")

if not login or not password:
    st.error("В secrets нет dataforseo.login или dataforseo.password")
    st.stop()

base_url = f"https://{host}"

run = st.button("Старт проверки", disabled=(uploaded is None))

if not run:
    st.stop()

# -----------------------
# Processing
# -----------------------
status_box = st.status("Подготовка...", expanded=True)
progress = st.progress(0.0)
stats = st.empty()

wb = load_workbook(BytesIO(uploaded.getvalue()))
tasks_meta, total = collect_tasks_from_workbook(wb)

if total == 0:
    status_box.update(label="В файле не найдено ни одной ссылки в колонке B", state="error")
    st.stop()

status_box.write(f"Найдено ссылок: **{total}** (все листы вместе).")
status_box.write("Этап 1/3: постановка задач (task_post)…")

session = requests.Session()
session.auth = (login, password)
session.headers.update({"Content-Type": "application/json"})

# task_id -> (sheet_name, row, url)
pending = {}
pending_ids = set()

BATCH_SIZE = 100
posted = 0
done = 0
errors = 0

def update_ui(stage_text: str, current_sheet: str | None = None):
    pct = done / total if total else 0.0
    progress.progress(min(pct, 1.0))
    sheet_text = f" | лист: `{current_sheet}`" if current_sheet else ""
    stats.markdown(
        f"**Этап:** {stage_text}{sheet_text}\n\n"
        f"**Готово:** {done}/{total}  \n"
        f"**Ошибок:** {errors}"
    )

# 1) POST batches
for i in range(0, total, BATCH_SIZE):
    batch_meta = tasks_meta[i:i + BATCH_SIZE]
    payload = []
    for (sheet_name, row, url) in batch_meta:
        payload.append({
            "keyword": build_site_query(url),
            "location_code": int(location_code),
            "language_code": str(language_code),
            "device": str(device),
            "depth": int(depth),
            "tag": f"{sheet_name}__{row}",
        })

    r = api_request(session, "POST", base_url + TASK_POST, json_body=payload, timeout=60)
    data = safe_json(r)

    if data.get("status_code") != 20000:
        # если упали на POST — сразу показываем и выходим
        status_box.update(label=f"Ошибка task_post: {data.get('status_code')} {data.get('status_message')}", state="error")
        st.code(data)
        st.stop()

    # Важно: задачи в ответе идут в том же порядке, что payload
    api_tasks = data.get("tasks") or []
    for idx, t in enumerate(api_tasks):
        tid = t.get("id")
        if not tid:
            continue
        sheet_name, row, url = batch_meta[idx]
        pending[tid] = (sheet_name, row, url)
        pending_ids.add(tid)

    posted += len(batch_meta)
    update_ui("Постановка задач (task_post)", batch_meta[-1][0])

status_box.write("Этап 2/3: ожидание готовности и получение результатов (tasks_ready + task_get)…")

# 2) Poll ready + GET results
last_poll = 0.0
POLL_EVERY = 3.0

while pending_ids:
    now = time.time()
    if now - last_poll < POLL_EVERY:
        time.sleep(0.5)
        continue
    last_poll = now

    update_ui("Получение результатов (tasks_ready/task_get)")

    r = api_request(session, "GET", base_url + TASKS_READY, timeout=60)
    data = safe_json(r)
    if data.get("status_code") != 20000:
        errors += 1
        continue

    ready = set()
    for t in data.get("tasks") or []:
        for res in (t.get("result") or []):
            rid = res.get("id")
            if rid:
                ready.add(rid)

    # берём только те, что относятся к нашему запуску
    to_process = list(pending_ids.intersection(ready))
    if not to_process:
        continue

    # Чтобы не застрелиться по лимитам — обрабатываем порциями
    to_process = to_process[:50]

    for tid in to_process:
        sheet_name, row, url = pending[tid]
        ws = wb[sheet_name]

        try:
            rr = api_request(session, "GET", base_url + TASK_GET_ADV.format(task_id=tid), timeout=120)
            res = safe_json(rr)

            if res.get("status_code") != 20000:
                ws.cell(row=row, column=4).value = "ERROR"
                errors += 1
            else:
                task = (res.get("tasks") or [{}])[0]
                if task.get("status_code") != 20000:
                    ws.cell(row=row, column=4).value = "ERROR"
                    errors += 1
                else:
                    result = (task.get("result") or [{}])[0]
                    items = result.get("items") or []
                    indexed = match_indexed(url, items)
                    # пишем в колонку D (Index)
                    ws.cell(row=row, column=4).value = bool(indexed)

            done += 1
            update_ui("Получение результатов (tasks_ready/task_get)", sheet_name)

        except Exception:
            ws.cell(row=row, column=4).value = "ERROR"
            done += 1
            errors += 1
            update_ui("Получение результатов (tasks_ready/task_get)", sheet_name)

        pending_ids.remove(tid)

status_box.write("Этап 3/3: формирование файла…")

# Save workbook to bytes
out = BytesIO()
wb.save(out)
out.seek(0)

status_box.update(label="Готово ✅", state="complete")
update_ui("Готово")

st.download_button(
    "Скачать результат (.xlsx)",
    data=out.getvalue(),
    file_name="indexation_checked.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
