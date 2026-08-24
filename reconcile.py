"""
Стол заказов — Скрипт ежедневной сверки v4.

Запускается через GitHub Actions (cron) раз в день в 09:00 МСК.

Что делает:
1. Скачивает CSV-фид наличия (drom-формат)
2. Читает из Google Sheets заявки со статусом «Ждёт деталь»
3. Сверяет category + part_name + position (+catalog_number) с фидом, исключая excluded_ids
4. Совпадения → статус «Найдено — связаться» + matched_at + дополняет matched_ids
5. Просроченные → статус «Просрочено»
6. По заявкам в статусе «Связались — ждём ответа» формирует напоминания
   по расписанию: 2-й день, потом каждые 7 дней (9, 16, 23, 30...)
7. Отправляет сводку в Telegram

Изменения v4:
- Добавлен 7-й статус «Связались — ждём ответа» — по нему сверка не идёт
- Новая скрытая колонка V `awaiting_since`. Скрипт записывает туда дату при первом
  обнаружении заявки в статусе AWAITING_REPLY и очищает при возврате в WAITING.
  Менеджер с этой колонкой не работает, она может быть скрыта через View → Hide column.
- Эскалирующие напоминания: 2 дня, потом каждые 7 (расписание в should_push_awaiting).
  Заявка не закрывается автоматически — пуши идут пока wait_until не сработает
  или менеджер не сменит статус.
- В итоговый блок сводки добавлен счётчик «Ждём ответа клиента»
- FINAL_STATUSES не изменился — AWAITING_REPLY не финальный статус
"""

import os
import csv
import io
import json
import urllib.request
import urllib.parse
from datetime import datetime, date, timedelta

# ─── Настройки из переменных окружения (GitHub Secrets) ───

FEED_URL = os.environ.get("FEED_URL", "https://baz-on.ru/export/c1326/2b944/razborangar-drom.csv")
TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]
GOOGLE_SHEETS_ID = os.environ["GOOGLE_SHEETS_ID"]
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
SHEETS_URL = os.environ.get("SHEETS_URL", "")

# Теги продавцов через запятую, например: "@ivanov,@petrov,@sidorov"
SELLER_TAGS = os.environ.get("SELLER_TAGS", "")

# ─── Колонки таблицы (0-indexed) ───

COL = {
    "id": 0,
    "created_at": 1,
    "client_name": 2,
    "contact_type": 3,
    "contact_nick": 4,
    "contact_phone": 5,
    "category": 6,
    "year": 7,
    "generation": 8,
    "restyling": 9,
    "vin": 10,
    "part_name": 11,
    "position": 12,
    "quantity": 13,
    "catalog_number": 14,
    "comment": 15,
    "wait_until": 16,
    "status": 17,
    "matched_at": 18,
    "excluded_ids": 19,
    "matched_ids": 20,  # Колонка U — скрипт дополняет, не перезаписывает
    "awaiting_since": 21,  # Колонка V — скрипт ставит при первом переходе в «Связались — ждём ответа», очищает при выходе
}

DATA_START_ROW = 4  # строки 1-3: заголовки, описания, пример

PRODUCT_URL_BASE = "https://razbor-angar.ru/p"

# ─── Статусы ───

STATUS_WAITING       = "Ждёт деталь"
STATUS_FOUND         = "Найдено — связаться"
STATUS_AWAITING_REPLY = "Связались — ждём ответа"   # v4: новый статус
STATUS_SOLD          = "Продано"
STATUS_DECLINED      = "Отказ"
STATUS_NO_CONTACT    = "Нет контакта"
STATUS_EXPIRED       = "Просрочено"

FINAL_STATUSES = {STATUS_SOLD, STATUS_DECLINED, STATUS_EXPIRED}

# Расписание напоминаний по статусу «Связались — ждём ответа»:
# первый пуш на 2-й день, потом каждые 7 дней (9, 16, 23, 30...).
# Чтобы изменить — поменять эти числа.
AWAITING_REPLY_FIRST_PUSH_DAY = 2
AWAITING_REPLY_REPEAT_EVERY_DAYS = 7


def should_push_awaiting(days_passed: int) -> bool:
    """День 2, 9, 16, 23, ... → True"""
    if days_passed < AWAITING_REPLY_FIRST_PUSH_DAY:
        return False
    if days_passed == AWAITING_REPLY_FIRST_PUSH_DAY:
        return True
    return (days_passed - AWAITING_REPLY_FIRST_PUSH_DAY) % AWAITING_REPLY_REPEAT_EVERY_DAYS == 0

# Динамическое заверение по дням недели
CLOSING_BY_WEEKDAY = {
    0: "Хорошего понедельника! ✌️",
    1: "Продуктивного вторника! ✌️",
    2: "Хорошей среды! ✌️",
    3: "Продуктивного четверга! ✌️",
    4: "Хорошей пятницы! ✌️",
    5: "Хорошей субботы! ✌️",
    6: "Хорошего воскресенья! ✌️",
}


# ═══════════════════════════════════════════
# Google Sheets API (через сервисный аккаунт)
# ═══════════════════════════════════════════

import google.auth
from google.oauth2 import service_account
from googleapiclient.discovery import build


def get_sheets_service():
    creds_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def read_orders(service):
    """Читает все строки из листа «Заявки» начиная с DATA_START_ROW."""
    result = service.spreadsheets().values().get(
        spreadsheetId=GOOGLE_SHEETS_ID,
        range=f"Заявки!A{DATA_START_ROW}:V",
    ).execute()
    return result.get("values", [])


def batch_update_cells(service, updates):
    """Пакетное обновление ячеек. updates = [{"row": 0, "col": 0, "value": "..."}]"""
    data = []
    for u in updates:
        col_letter = chr(ord("A") + u["col"])
        cell = f"Заявки!{col_letter}{DATA_START_ROW + u['row']}"
        data.append({"range": cell, "values": [[u["value"]]]})
    if not data:
        return
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=GOOGLE_SHEETS_ID,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()


# ═══════════════════════════════════════════
# CSV-фид наличия (drom-формат)
# ═══════════════════════════════════════════

def download_feed():
    """Скачивает CSV-фид (drom-формат), возвращает список товаров в наличии."""
    print(f"Downloading feed from {FEED_URL}...")
    resp = urllib.request.urlopen(FEED_URL)
    raw = resp.read().decode("cp1251")
    reader = csv.DictReader(io.StringIO(raw), delimiter=";", quotechar='"')

    items = []
    for row in reader:
        if row.get("Статус", "").strip() != "В наличии":
            continue

        item_id = row.get("Артикул", "").strip()

        # Собираем позицию из структурированных полей
        pos_parts = []
        pz = row.get("Перед/Зад", "").strip()
        lr = row.get("Лев/Прав", "").strip()
        if pz:
            pos_parts.append("передний" if "перед" in pz.lower() else "задний")
        if lr:
            pos_parts.append("левый" if "лев" in lr.lower() else "правый")
        feed_position = " ".join(pos_parts)

        # Краткое описание для Telegram: кузов, год
        desc_parts = []
        if row.get("Кузов", "").strip():
            desc_parts.append(row["Кузов"].strip())
        if row.get("Год", "").strip():
            desc_parts.append(row["Год"].strip())
        short_desc = ", ".join(desc_parts)

        brand = row.get("Марка", "").strip()
        model = row.get("Модель", "").strip()
        category = f"{brand} {model}".strip()

        price_raw = row.get("Цена", "").strip()
        try:
            price_val = float(price_raw) if price_raw else None
        except ValueError:
            price_val = None

        items.append({
            "id": item_id,
            "category": category,
            "name": row.get("Наименование", "").strip(),
            "price": price_raw,
            "price_val": price_val,
            "url": f"{PRODUCT_URL_BASE}{item_id}",
            "short_desc": short_desc,
            "position": feed_position,
            "vendor_codes": row.get("Номер", "").strip(),
            "brand": brand,
            "model": model,
            "year": row.get("Год", "").strip(),
        })

    print(f"Feed loaded: {len(items)} items available")
    return items


# ═══════════════════════════════════════════
# Матчинг
# ═══════════════════════════════════════════

def position_matches(order_position, feed_position):
    """
    Проверяет совместимость позиции заявки и товара из фида.
    Позиция из фида уже нормализована: 'передний правый', 'задний', '' и т.д.
    """
    if not order_position:
        return True
    if order_position in ("Любой", "Не важно"):
        return True
    if not feed_position:
        return True

    op = order_position.lower()
    fp = feed_position.lower()

    if op == fp:
        return True
    if op in ("передний", "задний"):
        return fp.startswith(op)
    if op in ("левый", "правый"):
        return fp.endswith(op)

    return False


def find_matches(order_category, order_part_name, order_position, order_catalog_number, excluded_ids_str, feed_items):
    """
    Ищет совпадения в фиде. Два режима:
    1. Если есть каталожный номер — ищет по нему среди всех vendor_codes товара
    2. Иначе — по category + name + position

    Исключает товары из excluded_ids.
    """
    excluded = set()
    if excluded_ids_str:
        for eid in excluded_ids_str.split(","):
            eid = eid.strip()
            if eid:
                excluded.add(eid)

    matches = []
    cat_num = order_catalog_number.strip().lower() if order_catalog_number else ""

    for item in feed_items:
        if item["id"] in excluded:
            continue

        # Режим 1: матч по каталожному номеру
        if cat_num:
            vendor_codes = [c.strip().lower() for c in item["vendor_codes"].split(",") if c.strip()]
            if cat_num in vendor_codes:
                if position_matches(order_position, item["position"]):
                    matches.append(item)
            continue

        # Режим 2: матч по category + name + position
        if (item["category"].lower() == order_category.lower() and
                item["name"].lower() == order_part_name.lower()):
            if position_matches(order_position, item["position"]):
                matches.append(item)

    return matches


def price_range(matches):
    """Возвращает строку диапазона цен или одиночную цену."""
    prices = [m["price_val"] for m in matches if m["price_val"] is not None]
    if not prices:
        return None
    lo = min(prices)
    hi = max(prices)
    if lo == hi:
        return f"{int(lo):,}".replace(",", " ") + " ₽"
    return f"{int(lo):,}".replace(",", " ") + " — " + f"{int(hi):,}".replace(",", " ") + " ₽"


def average_price(matches):
    """Возвращает среднюю цену по позициям или None."""
    prices = [m["price_val"] for m in matches if m["price_val"] is not None]
    if not prices:
        return None
    return sum(prices) / len(prices)


def merge_matched_ids(existing_str, new_ids):
    """Дополняет matched_ids новыми ID, не затирая старые."""
    existing = set()
    if existing_str:
        for eid in existing_str.split(","):
            e = eid.strip()
            if e:
                existing.add(e)
    merged = existing | set(new_ids)
    return ",".join(sorted(merged))


# ═══════════════════════════════════════════
# Telegram
# ═══════════════════════════════════════════

def send_telegram(text, silent=False):
    """silent=True — сообщение придёт без звука и вибрации (тихий день)."""
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = json.dumps({
        "chat_id": TELEGRAM_CHAT_ID,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
        "disable_notification": silent,
    }).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"})
    try:
        urllib.request.urlopen(req)
    except Exception as e:
        print(f"Telegram error: {e}")


def split_and_send_telegram(text, max_len=4000, silent=False):
    if len(text) <= max_len:
        send_telegram(text, silent=silent)
        return
    lines = text.split("\n")
    chunk = ""
    for line in lines:
        if len(chunk) + len(line) + 1 > max_len:
            send_telegram(chunk, silent=silent)
            chunk = ""
        chunk += line + "\n"
    if chunk.strip():
        send_telegram(chunk, silent=silent)


# ═══════════════════════════════════════════
# Основной процесс
# ═══════════════════════════════════════════

def get_cell(row, col_idx, default=""):
    """Безопасное получение значения ячейки."""
    if col_idx < len(row):
        return str(row[col_idx]).strip()
    return default


def parse_date_safe(s):
    """Парсит ISO-дату (YYYY-MM-DD), возвращает date или None."""
    if not s:
        return None
    try:
        return date.fromisoformat(s.strip())
    except (ValueError, AttributeError):
        return None


def main():
    today = date.today().isoformat()
    weekday = date.today().weekday()
    closing = CLOSING_BY_WEEKDAY[weekday]

    print(f"=== Сверка заявок: {today} ===")

    feed_items = download_feed()

    service = get_sheets_service()
    rows = read_orders(service)
    print(f"Orders loaded: {len(rows)} rows")

    updates = []
    clients = {}            # новые совпадения за сегодня (для шапки сообщения)
    stale_awaiting = []     # заявки «Связались — ждём ответа» старше порога

    # Счётчики по всему файлу
    total_active        = 0   # все не финальные статусы
    total_need_contact  = 0   # статус «Найдено — связаться»
    total_awaiting      = 0   # статус «Связались — ждём ответа»
    today_avg_sum       = 0.0 # сумма средних цен по новым совпадениям

    for i, row in enumerate(rows):
        status = get_cell(row, COL["status"])

        # ─── Счётчики по всему файлу ───
        if status and status not in FINAL_STATUSES:
            total_active += 1
        if status == STATUS_FOUND:
            total_need_contact += 1

        # ─── Обработка статуса «Связались — ждём ответа» ───
        if status == STATUS_AWAITING_REPLY:
            total_awaiting += 1

            awaiting_since_str = get_cell(row, COL["awaiting_since"])
            awaiting_dt = parse_date_safe(awaiting_since_str)

            if not awaiting_dt:
                # Первая сверка после того как менеджер поставил статус —
                # записываем сегодняшнюю дату как начало отсчёта
                updates.append({"row": i, "col": COL["awaiting_since"], "value": today})
            else:
                # Проверяем не пора ли пушнуть напоминание
                days_passed = (date.today() - awaiting_dt).days
                if should_push_awaiting(days_passed):
                    stale_awaiting.append({
                        "id": get_cell(row, COL["id"]),
                        "client_name": get_cell(row, COL["client_name"]),
                        "part_name": get_cell(row, COL["part_name"]),
                        "category": get_cell(row, COL["category"]),
                        "awaiting_since": awaiting_since_str,
                        "days_passed": days_passed,
                    })

        # Сверку проводим ТОЛЬКО для «Ждёт деталь».
        # «Связались — ждём ответа», «Найдено — связаться» и финальные статусы пропускаем.
        if status != STATUS_WAITING:
            continue

        # Если заявка вернулась в «Ждёт деталь» (например после excluded_ids) —
        # очищаем awaiting_since, чтобы при следующем переходе в AWAITING отсчёт начался заново
        if get_cell(row, COL["awaiting_since"]):
            updates.append({"row": i, "col": COL["awaiting_since"], "value": ""})

        order_id             = get_cell(row, COL["id"])
        client_name          = get_cell(row, COL["client_name"])
        contact_type         = get_cell(row, COL["contact_type"])
        contact_nick         = get_cell(row, COL["contact_nick"])
        contact_phone        = get_cell(row, COL["contact_phone"])
        category             = get_cell(row, COL["category"])
        year                 = get_cell(row, COL["year"])
        generation           = get_cell(row, COL["generation"])
        restyling            = get_cell(row, COL["restyling"])
        part_name            = get_cell(row, COL["part_name"])
        position             = get_cell(row, COL["position"])
        wait_until           = get_cell(row, COL["wait_until"])
        excluded_ids         = get_cell(row, COL["excluded_ids"])
        catalog_number       = get_cell(row, COL["catalog_number"])
        existing_matched_ids = get_cell(row, COL["matched_ids"])

        # Проверяем просрочку
        wait_dt = parse_date_safe(wait_until)
        if wait_dt and wait_dt < date.today():
            updates.append({"row": i, "col": COL["status"], "value": STATUS_EXPIRED})
            # Удалили из total_active, так как теперь финальный
            total_active -= 1
            continue

        matches = find_matches(category, part_name, position, catalog_number, excluded_ids, feed_items)
        if not matches:
            continue

        # Нашли — обновляем статус, matched_at, matched_ids
        updates.append({"row": i, "col": COL["status"], "value": STATUS_FOUND})
        updates.append({"row": i, "col": COL["matched_at"], "value": today})

        new_ids = [m["id"] for m in matches if m["id"]]
        merged = merge_matched_ids(existing_matched_ids, new_ids)
        updates.append({"row": i, "col": COL["matched_ids"], "value": merged})

        # Поднимаем счётчик «Нужно связаться» — заявка станет такой после апдейта
        total_need_contact += 1

        # Группировка для сообщения
        client_key = (client_name, contact_type, contact_nick, contact_phone, category, year, generation, restyling)
        if client_key not in clients:
            clients[client_key] = []

        part_display = part_name
        if position and position not in ("Любой", "Не важно"):
            part_display += f" ({position})"

        p_range = price_range(matches)

        part_block_lines = [f"<b>{order_id}</b> · {part_display}"]
        if p_range:
            part_block_lines.append(f"Диапазон цен: <b>{p_range}</b>")
        for m in matches:
            line = "• "
            if m["short_desc"]:
                line += f"{m['name']} ({m['short_desc']})"
            else:
                line += m["name"]
            if m["price"]:
                line += f" — {m['price']} ₽"
            part_block_lines.append(line)
            part_block_lines.append(f"  {m['url']}")

        clients[client_key].append("\n".join(part_block_lines))

        avg = average_price(matches)
        if avg:
            today_avg_sum += avg

    # ─── Применяем обновления пакетно ───
    if updates:
        print(f"Applying {len(updates)} updates...")
        batch_update_cells(service, updates)

    print(f"Results: {len(clients)} new matched, {total_need_contact} need contact, "
          f"{total_awaiting} awaiting reply, {len(stale_awaiting)} stale, {total_active} active")

    # ─── Формируем сводку в Telegram ───
    lines = []

    if clients:
        lines.append(f"<b>Всем привет!</b> Для {len(clients)} клиентов появились детали, они их ждут. Свяжитесь с ними сегодня! 🚗")

        for client_key, parts in clients.items():
            client_name, contact_type, contact_nick, contact_phone, category, year, generation, restyling = client_key

            lines.append("")
            lines.append("─ ─ ─")
            lines.append("")
            lines.append(f"<b>{client_name}</b>")

            contact_parts = []
            if contact_nick:
                contact_parts.append(f"{contact_type}: {contact_nick}")
            if contact_phone:
                contact_parts.append(f"Тел: {contact_phone}")
            if contact_parts:
                lines.append(" | ".join(contact_parts))

            car = category
            if year:
                car += f" {year}"
            if generation:
                car += f", {generation} поколение"
            if restyling:
                car += f", {restyling}"
            lines.append(car)
            lines.append("")

            for part_block in parts:
                lines.append(part_block)
                lines.append("")
    elif total_need_contact > 0:
        # Новых совпадений нет, но со вчера висят необработанные — это не тихий день
        lines.append(
            f"<b>Всем привет!</b> Новых совпадений сегодня нет, "
            f"но {total_need_contact} заявок ждут звонка со вчера. Разберём? 🚗"
        )
    else:
        lines.append("Сегодня новых совпадений нет.")

    # Выручка за сегодня (только по новым совпадениям)
    if today_avg_sum > 0:
        lines.append("─ ─ ─")
        lines.append("")
        revenue_str = f"{int(today_avg_sum):,}".replace(",", " ")
        lines.append(f"Потенциальная выручка при продаже сегодня: ~<b>{revenue_str} ₽</b>")
        lines.append("")

    # ─── Блок напоминаний по «Связались — ждём ответа» ───
    # Пуши приходят на 2-й, 9-й, 16-й день и далее каждые 7 дней.
    # Сортируем по убыванию: самые «зависшие» наверху, чтобы продавец видел их первыми.
    if stale_awaiting:
        stale_awaiting.sort(key=lambda x: x["days_passed"], reverse=True)
        lines.append("= = =")
        lines.append("")
        lines.append(f"⏰ <b>Дожимаем клиентов: {len(stale_awaiting)}</b>")
        lines.append("Висят в статусе «Связались — ждём ответа». Стоит дожать или закрыть.")
        lines.append("")
        for s in stale_awaiting:
            days = s["days_passed"]
            # Склонение «день»: 1 день, 2 дня, 5 дней, 21 день, 22 дня, 25 дней
            last_two = days % 100
            last_one = days % 10
            if 11 <= last_two <= 14:
                day_word = "дней"
            elif last_one == 1:
                day_word = "день"
            elif 2 <= last_one <= 4:
                day_word = "дня"
            else:
                day_word = "дней"
            age_label = f"уже {days} {day_word}"
            lines.append(
                f"• <b>{s['id']}</b> · {s['client_name']} · {s['category']} → {s['part_name']} "
                f"(<b>{age_label}</b>, с {s['awaiting_since']})"
            )
        lines.append("")

    # ─── Итоговый блок ───
    lines.append("= = =")
    lines.append("")
    lines.append(f"Всего актуальных заявок: {total_active}")
    lines.append(f"Нужно связаться с: <b>{total_need_contact}</b>")
    if total_awaiting > 0:
        lines.append(f"Ждём ответа клиента: {total_awaiting}")

    if SHEETS_URL:
        lines.append(f"\n<a href=\"{SHEETS_URL}\">Перейти в стол заказов</a>")

    # ─── Теги продавцов: только если есть работа на сегодня ───
    # needs_action складывается из двух источников:
    #   total_need_contact — статус «Найдено — связаться» (новые совпадения + вчерашние
    #                        необработанные, счётчик идёт по всему файлу)
    #   stale_awaiting     — клиенты в «Связались — ждём ответа», которых пора дожать
    # Если оба нуля — день тихий: без тегов и без звука.
    # Тег означает «тебе сейчас работать». Пинг вхолостую = бота замьютят.
    needs_action = total_need_contact + len(stale_awaiting)
    silent_day = (needs_action == 0)

    if SELLER_TAGS and not silent_day:
        tags = " ".join(t.strip() for t in SELLER_TAGS.split(",") if t.strip())
        lines.append(f"\n{tags}")

    lines.append(f"\n{closing}")

    full_message = "\n".join(lines)
    split_and_send_telegram(full_message, silent=silent_day)
    print(f"Done! needs_action={needs_action}, silent={silent_day}")


if __name__ == "__main__":
    main()
