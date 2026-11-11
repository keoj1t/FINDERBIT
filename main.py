import argparse
import hashlib
import html
import json
import logging
import os
import re
import sys
import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Optional
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
PROMPT_FILE = BASE_DIR / "prompt.txt"
SEARCH_PROMPT_FILE = BASE_DIR / "search_prompt.txt"
CRITERIA_FILE = BASE_DIR / "criteria.json"
OUTPUT_DIR = BASE_DIR / "output"
STATE_FILE = OUTPUT_DIR / "state.json"
# Ручной режим: текст поиска из Google AI Studio и (необязательно) готовый JSON.
NEWS_TEXT_FILE = BASE_DIR / "news.txt"
NEWS_JSON_FILE = BASE_DIR / "news.json"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(OUTPUT_DIR / "pipeline.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SEARCH_MODEL = os.getenv("GEMINI_SEARCH_MODEL", "gemini-3.5-flash-lite")
EXTRACT_MODEL = os.getenv("GEMINI_EXTRACT_MODEL", "gemini-3.1-flash-lite")
GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

BITRIX_WEBHOOK_URL = os.getenv("BITRIX_WEBHOOK_URL")
# Стадия лида. В Bitrix24 первая стадия воронки лидов имеет код NEW
# (в вашем портале она может называться «Неразобранное»). Если стадия
# называется иначе или код другой - задайте BITRIX_LEAD_STATUS в .env.
BITRIX_LEAD_STATUS = os.getenv("BITRIX_LEAD_STATUS", "NEW")
# Необязательно: код источника «ИИ-поиск», если он уже создан в CRM.
BITRIX_SOURCE_ID = os.getenv("BITRIX_SOURCE_ID")
# Необязательно: код пользовательского поля лида для БИН (например UF_CRM_BIN).
BITRIX_BIN_FIELD = os.getenv("BITRIX_BIN_FIELD")
# Необязательно: код поля БИН в реквизитах компании (для поиска дублей).
BITRIX_REQUISITE_BIN_FIELD = os.getenv("BITRIX_REQUISITE_BIN_FIELD")
# Необязательно: ID ответственного за создаваемые лиды.
BITRIX_ASSIGNED_BY_ID = os.getenv("BITRIX_ASSIGNED_BY_ID")
# Пауза между запросами: у облачного Bitrix24 лимит около 2 запросов в секунду.
BITRIX_PAUSE = float(os.getenv("BITRIX_PAUSE", "0.5"))

# Уведомления в Telegram (необязательно): токен бота и ID чата.
TELEGRAM_BOT_TOKEN = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip().strip("\"'")
TELEGRAM_CHAT_ID = (os.getenv("TELEGRAM_CHAT_ID") or "").strip().strip("\"'")
TELEGRAM_API = "https://api.telegram.org/bot{token}/{method}"
TELEGRAM_TIMEOUT = 60
TELEGRAM_MAX_LEADS_IN_MESSAGE = 15

GEMINI_TIMEOUT = 120
GEMINI_MAX_TOKENS = int(os.getenv("GEMINI_MAX_TOKENS", "16384"))
BITRIX_TIMEOUT = 15
BITRIX_RETRIES = 3
BIN_LENGTH = 12
NOT_SET = "Не установлено"

SOURCE_NAME = "ИИ-поиск"
SOURCE_CODE = "AI_SEARCH"
TITLE_PREFIX = "[AI-CLC]"
REVIEW_PREFIX = "[ПРОВЕРИТЬ ДУБЛЬ]"

# Статус каждого лида в CRM (колонка «Статус в CRM» в Excel).
CRM_NOT_SENT = "Не отправлялся в CRM"
CRM_CREATED = "Создан в CRM"
CRM_CREATED_REVIEW = "Создан в CRM (проверить дубль)"
CRM_STATUS_LABELS = {
    "created": CRM_CREATED,
    "review": CRM_CREATED_REVIEW,
    "updated": "Дубль: информация добавлена в карточку",
    "excluded": "Дубль: исключён",
    "known": "Уже обработан ранее",
    "error": "Ошибка передачи в CRM",
}

PRIORITY_ORDER = {"A": 0, "B": 1, "C": 2}
PRIORITY_COLORS = {"A": "C6EFCE", "B": "FFEB9C", "C": "E7E6E6"}
WORK_TYPE_LABELS = {
    "new_construction": "Новое строительство",
    "reconstruction": "Реконструкция",
    "repair": "Ремонт / капитальный ремонт",
    "landscaping": "Благоустройство",
}
REPAIR_SHEET_TYPES = ("reconstruction", "repair")

FONT_NAME = "Arial"
BODY_FONT = Font(name=FONT_NAME)
BOLD_FONT = Font(name=FONT_NAME, bold=True)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")
LINK_FONT = Font(name=FONT_NAME, color="0563C1", underline="single")
HEADER_FILL = PatternFill("solid", start_color="1F3864")
CELL_ALIGNMENT = Alignment(wrap_text=True, vertical="top", horizontal="left")

LEAD_COLUMNS = [
    ("project_name", "Наименование объекта", 40),
    ("city", "Город", 18),
    ("region", "Регион", 24),
    ("object_type", "Тип объекта", 24),
    ("work_type", "Вид работ", 24),
    ("company_name", "Заказчик", 32),
    ("bin", "БИН заказчика", 16),
    ("industry", "Сфера деятельности заказчика", 26),
    ("company_description", "Краткое описание заказчика", 40),
    ("website", "Сайт", 28),
    ("general_contractor", "Генеральный подрядчик", 32),
    ("finishing_subcontractor", "Субподрядчик по отделке", 32),
    ("stage", "Стадия", 28),
    ("timeline", "Сроки", 22),
    ("budget", "Бюджет", 20),
    ("clc_demand", "Потребность в продукции CLC", 40),
    ("contact_person", "Контактное лицо", 24),
    ("phone", "Телефон", 20),
    ("email", "Email", 26),
    ("source_link", "Источник", 40),
    ("publication_date", "Дата публикации", 18),
    ("recommendation", "Рекомендуемые действия", 45),
    ("priority", "Приоритет", 12),
    ("crm_status", "Статус в CRM", 30),
]
SOURCE_COLUMNS = [
    ("project_name", "Объект", 40),
    ("source_link", "Ссылка на источник", 60),
    ("publication_date", "Дата публикации", 18),
]
LEAD_FIELDS = [key for key, _, _ in LEAD_COLUMNS]
ANALYTICS_FIELDS = ("project_name", "fact", "impact", "action")

# Поля и подписи для комментария в карточке лида.
COMMENT_FIELDS = [
    ("project_name", "Объект"),
    ("object_type", "Тип объекта"),
    ("work_type", "Вид работ"),
    ("city", "Город"),
    ("region", "Регион"),
    ("industry", "Сфера деятельности заказчика"),
    ("company_description", "О заказчике"),
    ("bin", "БИН"),
    ("website", "Сайт"),
    ("general_contractor", "Генподрядчик"),
    ("finishing_subcontractor", "Субподрядчик по отделке"),
    ("stage", "Стадия"),
    ("timeline", "Сроки"),
    ("budget", "Бюджет"),
    ("clc_demand", "Потребность в продукции CLC"),
    ("contact_person", "Контактное лицо"),
    ("phone", "Телефон"),
    ("email", "Email"),
    ("publication_date", "Дата публикации"),
    ("priority", "Приоритет"),
    ("recommendation", "Рекомендация CLC"),
    ("source_link", "Источник"),
]

LEGAL_FORMS = {
    "тоо", "ао", "кгу", "ргу", "гу", "кгп", "ргп", "ип", "оао", "зао", "ооо",
    "llp", "jsc", "тов",
}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# --------------------------------------------------------------------------
# Общие вспомогательные функции
# --------------------------------------------------------------------------
def is_set(value: Optional[str]) -> bool:
    """Возвращает True, если значение заполнено и не равно заглушке."""
    return bool(value) and value != NOT_SET


def as_text(value: Any) -> str:
    """Приводит значение к непустой строке или к заглушке NOT_SET."""
    if value is None:
        return NOT_SET
    text = str(value).strip()
    return text or NOT_SET


def as_list(value: Any) -> list:
    """Возвращает значение, если это список, иначе пустой список."""
    return value if isinstance(value, list) else []


def is_url(value: str) -> bool:
    return value.startswith(("http://", "https://")) and len(value) < 2000


def digits(value: str) -> str:
    return re.sub(r"\D", "", value)


def split_values(value: str) -> list[str]:
    """Делит строку с несколькими телефонами/email на отдельные значения."""
    return [part.strip() for part in re.split(r"[;,\n]+", value) if part.strip()]


def website_host(value: str) -> str:
    url = value if "://" in value else f"https://{value}"
    host = urlparse(url).netloc.lower()
    return host[4:] if host.startswith("www.") else host


def normalize_name(value: str) -> str:
    """Название организации без кавычек и организационно-правовой формы."""
    text = value.lower().replace("ё", "е")
    text = re.sub(r"[«»\"'“”„().,;:–—-]", " ", text)
    return " ".join(word for word in text.split() if word not in LEGAL_FORMS)


def lead_fingerprint(lead: dict) -> str:
    """Устойчивый ключ «заказчик + объект» для защиты от повторной обработки."""
    owner = (
        lead["bin"] if is_set(lead["bin"]) else normalize_name(lead["company_name"])
    )
    what = (
        normalize_name(lead["project_name"])
        if is_set(lead["project_name"])
        else lead["source_link"]
    )
    return hashlib.sha1(f"{owner}|{what}".encode("utf-8")).hexdigest()[:16]


def load_text(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    return path.read_text(encoding="utf-8")


# --------------------------------------------------------------------------
# Статистика запуска и состояние
# --------------------------------------------------------------------------
@dataclass
class RunStats:
    found: int = 0           # сколько компаний/объектов вернула модель
    matched: int = 0         # сколько из них прошли проверку и соответствуют критериям
    created: int = 0         # создано лидов
    flagged: int = 0         # из созданных - с пометкой «проверить дубль»
    duplicates: int = 0      # исключено дублей (лид не создавался)
    updated: int = 0         # из дублей - дополнена существующая карточка
    unverified_removed: int = 0  # значения, которых нет в тексте источника
    errors: int = 0
    error_details: list = field(default_factory=list)

    def add_error(self, message: str) -> None:
        self.errors += 1
        self.error_details.append(message)
        logger.error(message)


def load_state() -> dict:
    """Отпечатки уже обработанных лидов: защита от повторной отправки."""
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_state(state: dict) -> None:
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def format_run_report(stats: RunStats, push: bool, dry_run: bool) -> str:
    mode = ""
    if dry_run:
        mode = " (тестовый режим: CRM не изменялась)"
    elif not push:
        mode = " (отправка в Bitrix24 отключена)"

    created = f"Создано лидов: {stats.created}"
    if stats.flagged:
        created += f" (из них с пометкой «проверить дубль»: {stats.flagged})"
    duplicates = f"Исключено дублей: {stats.duplicates}"
    if stats.updated:
        duplicates += f" (в существующие карточки добавлена новая информация: {stats.updated})"

    lines = [
        f"Отчёт о выгрузке за {date.today().isoformat()}{mode}",
        f"Найдено компаний/объектов: {stats.found}",
        f"Соответствуют критериям: {stats.matched}",
        created,
        duplicates,
        f"Ошибок при поиске или передаче данных: {stats.errors}",
    ]
    if stats.unverified_removed:
        lines.append(
            "Удалено непроверенных значений (их нет в тексте источника): "
            f"{stats.unverified_removed}"
        )
    for detail in stats.error_details:
        lines.append(f"  - {detail}")
    return "\n".join(lines)


def save_run_report(stats: RunStats, push: bool, dry_run: bool) -> None:
    text = format_run_report(stats, push, dry_run)
    logger.info("\n%s", text)
    path = OUTPUT_DIR / f"run_report_{date.today().isoformat()}.txt"
    with path.open("a", encoding="utf-8") as handle:
        handle.write(f"[{datetime.now():%H:%M:%S}]\n{text}\n\n")


# --------------------------------------------------------------------------
# Gemini
# --------------------------------------------------------------------------
class GeminiAccessError(Exception):
    """Нет доступа к Gemini API: квота, оплата, права или не задан ключ."""


def is_access_problem(status: int, body: str) -> bool:
    """429 (квота), 402/403 (оплата, права), 400 FAILED_PRECONDITION (нужна оплата)."""
    if status in (402, 403, 429):
        return True
    lowered = body.lower()
    return status == 400 and ("failed_precondition" in lowered or "billing" in lowered)


def call_gemini(
    model: str,
    prompt: str,
    system_prompt: Optional[str] = None,
    use_search: bool = False,
    json_mode: bool = False,
) -> Optional[dict]:
    """Вызывает Gemini API и возвращает первый вариант ответа (candidate)."""
    if not GEMINI_API_KEY:
        raise GeminiAccessError("GEMINI_API_KEY не задан")
    url = GEMINI_URL.format(model=model)
    headers = {
        "x-goog-api-key": GEMINI_API_KEY,
        "Content-Type": "application/json",
    }
    payload: dict[str, Any] = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": GEMINI_MAX_TOKENS},
    }
    if system_prompt:
        payload["systemInstruction"] = {"parts": [{"text": system_prompt}]}
    if use_search:
        payload["tools"] = [{"google_search": {}}]
    if json_mode:
        payload["generationConfig"]["responseMimeType"] = "application/json"

    try:
        response = requests.post(
            url, headers=headers, json=payload, timeout=GEMINI_TIMEOUT
        )
        response.raise_for_status()
        candidate = response.json()["candidates"][0]
    except requests.HTTPError as exc:
        if is_access_problem(exc.response.status_code, exc.response.text):
            raise GeminiAccessError(
                f"HTTP {exc.response.status_code} для модели {model}"
            ) from exc
        logger.error(
            "Gemini returned HTTP %s for model '%s': %s",
            exc.response.status_code,
            model,
            exc.response.text[:300],
        )
        return None
    except requests.RequestException as exc:
        logger.error("Gemini request failed for model '%s': %s", model, exc)
        return None
    except (KeyError, IndexError, ValueError) as exc:
        logger.error("Unexpected Gemini response format: %s", exc)
        return None

    if candidate.get("finishReason") == "MAX_TOKENS":
        logger.warning("Gemini answer was cut off: the token limit was reached")
    return candidate


def candidate_text(candidate: dict) -> str:
    """Склеивает текстовые части ответа, пропуская служебные части с рассуждениями."""
    parts = candidate.get("content", {}).get("parts", [])
    return "".join(
        part.get("text", "") for part in parts if not part.get("thought")
    ).strip()


def resolve_url(uri: str) -> str:
    """Google Search отдает ссылки-перенаправления; разворачиваем в прямую ссылку."""
    if "vertexaisearch.cloud.google.com" not in uri:
        return uri
    try:
        with requests.get(uri, allow_redirects=True, timeout=15, stream=True) as resp:
            return resp.url
    except requests.RequestException as exc:
        logger.warning("Could not resolve redirect %s: %s", uri[:80], exc)
        return uri


def grounding_sources(candidate: dict) -> list[str]:
    """Возвращает прямые ссылки на источники, которые использовал Google Search."""
    chunks = candidate.get("groundingMetadata", {}).get("groundingChunks", [])
    sources = []
    for chunk in chunks:
        web = chunk.get("web") or {}
        if web.get("uri"):
            sources.append(f"{web.get('title', 'источник')}: {resolve_url(web['uri'])}")
    return list(dict.fromkeys(sources))


def render_search_prompt() -> str:
    """Подставляет критерии из criteria.json в шаблон search_prompt.txt."""
    template = load_text(SEARCH_PROMPT_FILE)
    criteria = json.loads(load_text(CRITERIA_FILE))

    def bullets(items: list) -> str:
        return "\n".join(f"- {item}" for item in items)

    values = {
        "{today}": date.today().isoformat(),
        "{days}": str(criteria["days"]),
        "{max_projects}": str(criteria["max_projects"]),
        "{regions}": ", ".join(criteria["regions"]),
        "{object_types}": bullets(criteria["object_types"]),
        "{work_types}": bullets(criteria["work_types"]),
        "{keywords}": ", ".join(criteria["keywords"]),
        "{exclude}": ", ".join(criteria["exclude_keywords"]),
        "{sources}": bullets(criteria["sources"]),
    }
    for placeholder, value in values.items():
        template = template.replace(placeholder, value)
    return template


def search_projects_text(search_prompt: str) -> Optional[str]:
    """Первый шаг: Gemini с Google Search ищет проекты и возвращает текст с источниками."""
    logger.info("Searching the web with model '%s'", SEARCH_MODEL)
    candidate = call_gemini(SEARCH_MODEL, search_prompt, use_search=True)
    if candidate is None:
        return None

    text = candidate_text(candidate)
    if not text:
        logger.error("Gemini returned an empty answer for the search step")
        return None

    sources = grounding_sources(candidate)
    if sources:
        text += "\n\nИсточники, найденные поиском:\n" + "\n".join(sources)
    return text


def parse_json_object(content: str) -> Optional[dict]:
    """Достает JSON-объект из ответа модели, даже если вокруг есть лишний текст."""
    start = content.find("{")
    end = content.rfind("}")
    if start == -1 or end <= start:
        return None
    try:
        parsed = json.loads(content[start : end + 1])
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def extract_report_via_gemini(raw_text: str, system_prompt: str) -> Optional[dict]:
    """Второй шаг: Gemini превращает найденный текст в JSON-отчет по схеме из промпта."""
    prompt = (
        "Extract construction leads for CLC Company from the text below and "
        "return the report as a JSON object that follows the schema from the "
        f"system prompt.\n\n{raw_text}"
    )
    candidate = call_gemini(
        EXTRACT_MODEL, prompt, system_prompt=system_prompt, json_mode=True
    )
    if candidate is None:
        return None

    report = parse_json_object(candidate_text(candidate))
    if report is None:
        logger.error("Model answer does not contain a valid JSON object")
    return report


# --------------------------------------------------------------------------
# Нормализация и проверка данных
# --------------------------------------------------------------------------
def clean_bin(value: str) -> str:
    if not is_set(value):
        return NOT_SET
    code = digits(value)
    return code if len(code) == BIN_LENGTH else NOT_SET


def clean_phones(value: str) -> str:
    if not is_set(value):
        return NOT_SET
    parts = [p for p in split_values(value) if len(digits(p)) >= 6]
    return "; ".join(parts) or NOT_SET


def clean_emails(value: str) -> str:
    if not is_set(value):
        return NOT_SET
    parts = [p for p in split_values(value) if EMAIL_RE.match(p)]
    return "; ".join(parts) or NOT_SET


def clean_website(value: str) -> str:
    if not is_set(value):
        return NOT_SET
    url = value if value.startswith(("http://", "https://")) else f"https://{value}"
    return url if "." in website_host(url) else NOT_SET


def normalize_lead(raw: dict) -> dict:
    """Приводит лид к единому виду: все поля есть, значения корректны."""
    lead = {key: as_text(raw.get(key)) for key in LEAD_FIELDS}

    lead["priority"] = lead["priority"].upper()
    if lead["priority"] not in PRIORITY_ORDER:
        lead["priority"] = "C"

    lead["work_type"] = lead["work_type"].lower()
    if lead["work_type"] not in WORK_TYPE_LABELS:
        lead["work_type"] = NOT_SET

    lead["bin"] = clean_bin(lead["bin"])
    lead["phone"] = clean_phones(lead["phone"])
    lead["email"] = clean_emails(lead["email"])
    lead["website"] = clean_website(lead["website"])
    if not is_url(lead["source_link"]):
        lead["source_link"] = NOT_SET
    return lead


def verify_lead(lead: dict, source_text: str, stats: RunStats) -> None:
    """Удаляет БИН, телефон, email, сайт и ссылку, которых нет в тексте источника.

    ТЗ: система не должна придумывать отсутствующие контактные данные.
    """
    lowered = source_text.lower()
    only_digits = digits(source_text)

    def check(key: str, found) -> None:
        if is_set(lead[key]) and not found(lead[key]):
            logger.warning("'%s' not found in source text, removed: %s", key, lead[key])
            lead[key] = NOT_SET
            stats.unverified_removed += 1

    check("bin", lambda v: v in only_digits)
    check(
        "phone",
        lambda v: all(digits(p)[-10:] in only_digits for p in split_values(v)),
    )
    check("email", lambda v: all(p.lower() in lowered for p in split_values(v)))
    check("website", lambda v: website_host(v) in lowered)
    check("source_link", lambda v: v.rstrip("/").lower() in lowered)


def normalize_report(raw: dict) -> dict:
    """Проверяет и нормализует ответ модели."""
    leads = [
        normalize_lead(item)
        for item in as_list(raw.get("leads"))
        if isinstance(item, dict)
    ]
    # Лид без названия объекта и без заказчика не несёт информации.
    leads = [
        lead
        for lead in leads
        if is_set(lead["company_name"]) or is_set(lead["project_name"])
    ]
    leads.sort(key=lambda lead: PRIORITY_ORDER[lead["priority"]])

    raw_summary = raw.get("summary")
    if not isinstance(raw_summary, dict):
        raw_summary = {}

    analytics = [
        {key: as_text(item.get(key)) for key in ANALYTICS_FIELDS}
        for item in as_list(raw_summary.get("analytics"))
        if isinstance(item, dict)
    ]

    return {
        "report_date": date.today().isoformat(),
        "leads": leads,
        "summary": {
            "top_projects": [
                as_text(name) for name in as_list(raw_summary.get("top_projects"))
            ][:4],
            "priority_actions": [
                as_text(action)
                for action in as_list(raw_summary.get("priority_actions"))
            ],
            "analytics": analytics,
            "tech_news": as_text(raw_summary.get("tech_news")),
        },
    }


def dedupe_batch(leads: list[dict], stats: RunStats) -> list[dict]:
    """Убирает повторы внутри одной выгрузки (остаётся лид с высшим приоритетом)."""
    seen: set[str] = set()
    unique = []
    for lead in leads:
        fingerprint = lead_fingerprint(lead)
        if fingerprint in seen:
            stats.duplicates += 1
            logger.info("Duplicate inside the batch: %s", lead["project_name"])
            continue
        seen.add(fingerprint)
        unique.append(lead)
    return unique


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def put(
    ws: Worksheet,
    row: int,
    column: int,
    value: Any,
    font: Font = BODY_FONT,
    formula: bool = False,
):
    """Записывает значение в ячейку. Текст, начинающийся с '=', не считается формулой."""
    cell = ws.cell(row=row, column=column, value=value)
    if isinstance(value, str) and value.startswith("=") and not formula:
        cell.data_type = "s"
    cell.font = font
    cell.alignment = CELL_ALIGNMENT
    return cell


def fill_table_sheet(ws: Worksheet, columns: list, rows: list[dict]) -> None:
    """Заполняет лист таблицей с заголовками, фильтром и закрепленной шапкой."""
    for index, (_, header, width) in enumerate(columns, start=1):
        cell = put(ws, 1, index, header, font=HEADER_FONT)
        cell.fill = HEADER_FILL
        ws.column_dimensions[get_column_letter(index)].width = width

    for row_index, row in enumerate(rows, start=2):
        for col_index, (key, _, _) in enumerate(columns, start=1):
            value = row[key]
            cell = put(ws, row_index, col_index, value)
            if key == "source_link" and is_url(value):
                cell.hyperlink = value
                cell.font = LINK_FONT
            elif key == "priority":
                cell.fill = PatternFill("solid", start_color=PRIORITY_COLORS[value])

    last_column = get_column_letter(len(columns))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{last_column}{max(len(rows) + 1, 2)}"


def to_display_rows(leads: list[dict]) -> list[dict]:
    """Подставляет русские названия видов работ для Excel."""
    return [
        {**lead, "work_type": WORK_TYPE_LABELS.get(lead["work_type"], NOT_SET)}
        for lead in leads
    ]


def write_heading(ws: Worksheet, row: int, text: str) -> int:
    put(ws, row, 1, text, font=BOLD_FONT)
    return row + 1


def write_pair(
    ws: Worksheet, row: int, label: str, value: Any, formula: bool = False
) -> int:
    put(ws, row, 1, label)
    put(ws, row, 2, value, formula=formula)
    return row + 1


def write_summary_sheet(ws: Worksheet, report: dict, last_row: int) -> None:
    """Заполняет лист «Сводка». Счетчики считаются формулами по листу «Новые лиды»."""
    leads = report["leads"]
    summary = report["summary"]
    sheet_ref = "'Новые лиды'"

    def column_range(key: str) -> str:
        letter = get_column_letter(LEAD_FIELDS.index(key) + 1)
        return f"{sheet_ref}!${letter}$2:${letter}${last_row}"

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 90

    put(ws, 1, 1, f"Сводка за {report['report_date']}", font=TITLE_FONT)
    row = 3

    if not leads:
        row = write_pair(
            ws,
            row,
            "Примечание",
            "Достоверных новых лидов в предоставленном тексте не найдено.",
        )
        row += 1

    row = write_heading(ws, row, "Общее количество лидов")
    row = write_pair(
        ws, row, "Всего", f"=COUNTA({column_range('project_name')})", formula=True
    )
    row += 1

    row = write_heading(ws, row, "Лиды по приоритетам")
    for priority in PRIORITY_ORDER:
        formula = f'=COUNTIF({column_range("priority")},"{priority}")'
        row = write_pair(ws, row, f"Категория {priority}", formula, formula=True)
    row += 1

    for title, key in (
        ("Распределение по регионам", "region"),
        ("Распределение по типам объектов", "object_type"),
    ):
        row = write_heading(ws, row, title)
        for name, _ in Counter(lead[key] for lead in leads).most_common():
            formula = f"=COUNTIF({column_range(key)},$A{row})"
            row = write_pair(ws, row, name, formula, formula=True)
        row += 1

    top_projects = summary["top_projects"] or [
        lead["project_name"] for lead in leads[:4]
    ]
    row = write_heading(ws, row, "Наиболее перспективные объекты")
    for name in top_projects:
        put(ws, row, 1, name)
        row += 1
    row += 1

    row = write_heading(ws, row, "Первоочередные действия отдела продаж")
    for number, action in enumerate(summary["priority_actions"], start=1):
        row = write_pair(ws, row, f"Действие {number}", action)
    row += 1

    row = write_heading(ws, row, "Аналитика по перспективным проектам")
    for item in summary["analytics"]:
        put(ws, row, 1, item["project_name"], font=BOLD_FONT)
        row += 1
        row = write_pair(ws, row, "Факт", item["fact"])
        row = write_pair(ws, row, "Влияние на CLC", item["impact"])
        row = write_pair(ws, row, "Действие", item["action"])
        row += 1

    row = write_heading(ws, row, "Новость о технологиях, ИИ или автоматизации")
    write_pair(ws, row, "Новость", summary["tech_news"])


def build_excel(report: dict, path: Path) -> None:
    """Собирает Excel-файл с четырьмя листами и сохраняет его."""
    leads = report["leads"]
    repairs = [lead for lead in leads if lead["work_type"] in REPAIR_SHEET_TYPES]

    workbook = Workbook()
    leads_sheet = workbook.active
    leads_sheet.title = "Новые лиды"
    repairs_sheet = workbook.create_sheet("Реконструкции и ремонты")
    summary_sheet = workbook.create_sheet("Сводка")
    sources_sheet = workbook.create_sheet("Источники")

    fill_table_sheet(leads_sheet, LEAD_COLUMNS, to_display_rows(leads))
    fill_table_sheet(repairs_sheet, LEAD_COLUMNS, to_display_rows(repairs))
    write_summary_sheet(summary_sheet, report, last_row=max(len(leads) + 1, 2))
    fill_table_sheet(sources_sheet, SOURCE_COLUMNS, leads)

    workbook.save(path)


def save_search_text(text: str) -> None:
    """Сохраняет текст, найденный поиском, чтобы можно было проверить источники."""
    path = OUTPUT_DIR / f"search_{date.today().isoformat()}.txt"
    path.write_text(text, encoding="utf-8")
    logger.info("Search text saved: %s", path)


def save_outputs(report: dict, retry: bool = False) -> Path:
    """Сохраняет JSON и Excel в папку output и возвращает путь к Excel.

    При повторной обработке (--from-json) JSON не перезаписывается,
    Excel сохраняется отдельным файлом report_retry_<дата>.xlsx.
    """
    stamp = report["report_date"]

    if not retry:
        json_path = OUTPUT_DIR / f"leads_{stamp}.json"
        json_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        logger.info("JSON saved: %s", json_path)

    excel_path = OUTPUT_DIR / (f"report_retry_{stamp}.xlsx" if retry else f"report_{stamp}.xlsx")
    build_excel(report, excel_path)
    logger.info("Excel saved: %s", excel_path)
    return excel_path


def save_failed(leads: list[dict]) -> None:
    """Лиды, которые не удалось передать в CRM. Повтор: --from-json <файл>."""
    path = OUTPUT_DIR / f"failed_{date.today().isoformat()}.json"
    path.write_text(
        json.dumps({"leads": leads, "summary": {}}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.warning(
        "%d lead(s) were not delivered. Retry with: --from-json %s", len(leads), path
    )


# --------------------------------------------------------------------------
# Telegram
# --------------------------------------------------------------------------
def telegram_enabled() -> bool:
    return bool(TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID)


def telegram_call(method: str, data: dict, files: Optional[dict] = None) -> bool:
    """Вызывает метод Telegram Bot API. Токен в логи не попадает."""
    url = TELEGRAM_API.format(token=TELEGRAM_BOT_TOKEN, method=method)
    try:
        response = requests.post(url, data=data, files=files, timeout=TELEGRAM_TIMEOUT)
        result = response.json()
    except (requests.RequestException, ValueError) as exc:
        logger.warning("Telegram %s failed: %s", method, type(exc).__name__)
        return False
    if not isinstance(result, dict) or not result.get("ok"):
        description = result.get("description") if isinstance(result, dict) else ""
        logger.warning(
            "Telegram %s rejected: %s", method, description or response.status_code
        )
        return False
    return True


def send_telegram_message(text: str) -> bool:
    return telegram_call(
        "sendMessage",
        {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": text[:4000],
            "disable_web_page_preview": "true",
        },
    )


def send_telegram_file(path: Path, caption: str) -> bool:
    with path.open("rb") as handle:
        return telegram_call(
            "sendDocument",
            {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
            files={"document": (path.name, handle)},
        )


def format_telegram_summary(stats: RunStats, leads: list[dict]) -> str:
    created = [
        lead for lead in leads
        if lead.get("crm_status") in (CRM_CREATED, CRM_CREATED_REVIEW)
    ]
    if stats.created:
        lines = [f"✅ CLC: в CRM Bitrix24 добавлено лидов: {stats.created}"]
    else:
        lines = ["ℹ️ CLC: новых лидов в CRM не добавлено"]
    if stats.flagged:
        lines.append(f"Из них с пометкой «проверить дубль»: {stats.flagged}")

    duplicates = f"Дублей исключено: {stats.duplicates}"
    if stats.updated:
        duplicates += f" (в существующие карточки добавлена информация: {stats.updated})"
    lines += [
        f"Найдено: {stats.found}, подходят критериям: {stats.matched}",
        duplicates,
    ]
    if stats.errors:
        lines.append(f"⚠️ Ошибок: {stats.errors}")
        lines += [f"  - {detail}" for detail in stats.error_details[:5]]

    if created:
        lines += ["", "Новые лиды:"]
        for number, lead in enumerate(created[:TELEGRAM_MAX_LEADS_IN_MESSAGE], start=1):
            place = f" ({lead['city']})" if is_set(lead["city"]) else ""
            lines.append(
                f"{number}. [{lead['priority']}] {lead['company_name']} — "
                f"{lead['project_name']}{place}"
            )
        hidden = len(created) - TELEGRAM_MAX_LEADS_IN_MESSAGE
        if hidden > 0:
            lines.append(f"… и ещё {hidden}, полный список в Excel")
    lines += ["", f"Выгрузка за {date.today().isoformat()}"]
    return "\n".join(lines)


def notify_telegram(
    stats: RunStats,
    report: Optional[dict],
    excel_path: Optional[Path],
    args: argparse.Namespace,
) -> None:
    """Сводка в Telegram после работы с CRM; Excel - если есть добавленные лиды."""
    if args.dry_run or args.skip_bitrix:
        return  # CRM не менялась, сообщать не о чем
    if not telegram_enabled():
        logger.info("Telegram is not configured, notification skipped")
        return

    leads = report["leads"] if report else []
    if send_telegram_message(format_telegram_summary(stats, leads)):
        logger.info("Telegram summary sent")
    if excel_path and excel_path.exists() and (stats.created or stats.updated):
        caption = f"Отчёт CLC за {date.today().isoformat()}: добавлено лидов {stats.created}"
        if send_telegram_file(excel_path, caption):
            logger.info("Telegram Excel sent: %s", excel_path.name)


# --------------------------------------------------------------------------
# Bitrix24
# --------------------------------------------------------------------------
class BitrixError(Exception):
    """Ошибка вызова REST API Bitrix24."""


_last_bitrix_call = 0.0
_source_id_cache: Optional[str] = None


def call_bitrix(method: str, payload: dict) -> dict[str, Any]:
    """Вызывает метод REST API Bitrix24. При ошибке бросает BitrixError."""
    global _last_bitrix_call
    url = f"{BITRIX_WEBHOOK_URL.rstrip('/')}/{method}.json"
    last_error = ""

    for attempt in range(1, BITRIX_RETRIES + 1):
        pause = BITRIX_PAUSE - (time.monotonic() - _last_bitrix_call)
        if pause > 0:
            time.sleep(pause)
        try:
            response = requests.post(url, json=payload, timeout=BITRIX_TIMEOUT)
        except requests.RequestException as exc:
            _last_bitrix_call = time.monotonic()
            last_error = type(exc).__name__  # текст ошибки содержит секрет вебхука
            time.sleep(attempt)
            continue
        _last_bitrix_call = time.monotonic()

        try:
            data = response.json()
        except ValueError:
            raise BitrixError(
                f"{method}: HTTP {response.status_code}, ответ не в формате JSON"
            )

        if isinstance(data, dict) and "error" in data:
            if data["error"] == "QUERY_LIMIT_EXCEEDED":
                last_error = "превышен лимит запросов"
                time.sleep(2 * attempt)
                continue
            raise BitrixError(
                f"{method}: {data.get('error_description', data['error'])}"
            )
        if not response.ok:
            raise BitrixError(f"{method}: HTTP {response.status_code}")
        return data

    raise BitrixError(f"{method}: не удалось выполнить запрос ({last_error})")


ENTITY_API = {
    "lead": ("crm.lead.list", ["ID", "TITLE", "COMMENTS", "WEB"]),
    "company": ("crm.company.list", ["ID", "TITLE", "COMMENTS", "WEB"]),
    "contact": ("crm.contact.list", ["ID", "NAME", "LAST_NAME"]),
}
ENTITY_TYPE_COMPANY = "4"


def list_entities(kind: str, filter_fields: dict) -> list[dict]:
    method, select = ENTITY_API[kind]
    data = call_bitrix(method, {"filter": filter_fields, "select": select})
    result = data.get("result")
    return result if isinstance(result, list) else []


@dataclass
class Match:
    level: str          # "exact" - точное совпадение, "partial" - частичное
    kind: str           # lead / company / contact
    entity_id: str
    title: str
    same_project: bool  # этот объект уже записан в найденной карточке
    reason: str


def entity_title(row: dict) -> str:
    return row.get("TITLE") or f"{row.get('NAME', '')} {row.get('LAST_NAME', '')}".strip()


def is_same_project(row: dict, lead: dict) -> bool:
    text = html.unescape(f"{row.get('TITLE', '')} {row.get('COMMENTS', '')}").lower()
    if is_set(lead["project_name"]) and lead["project_name"].lower() in text:
        return True
    return is_url(lead["source_link"]) and lead["source_link"].lower() in text


def make_match(level: str, kind: str, rows: list[dict], lead: dict, reason: str) -> Match:
    row = next((r for r in rows if is_same_project(r, lead)), rows[0])
    return Match(
        level, kind, str(row["ID"]), entity_title(row), is_same_project(row, lead), reason
    )


def lead_company_part(title: str) -> str:
    """Из заголовка «[AI-CLC] Заказчик — Объект» берёт название заказчика."""
    for prefix in (REVIEW_PREFIX, TITLE_PREFIX):
        title = title.replace(prefix, "")
    return title.split(" — ")[0].strip()


def find_existing(lead: dict) -> Optional[Match]:
    """Ищет заказчика в CRM среди лидов, компаний и контактов.

    Признаки: БИН, телефон, email, сайт, название организации.
    Точное совпадение - exact; совпадение только части названия - partial.
    """
    company, bin_code = lead["company_name"], lead["bin"]

    # 1. БИН: в комментариях лидов и компаний (и в реквизитах, если задано поле).
    if is_set(bin_code):
        for kind in ("lead", "company"):
            # Bitrix молча игнорирует неподдерживаемый фильтр и отдаёт все записи,
            # поэтому БИН проверяем в самих найденных карточках.
            rows = [
                r for r in list_entities(kind, {"%COMMENTS": bin_code})
                if bin_code in digits(html.unescape(r.get("COMMENTS") or ""))
            ]
            if rows:
                return make_match("exact", kind, rows, lead, "БИН")
        if BITRIX_REQUISITE_BIN_FIELD:
            data = call_bitrix(
                "crm.requisite.list",
                {
                    "filter": {BITRIX_REQUISITE_BIN_FIELD: bin_code},
                    "select": ["ID", "ENTITY_ID", "ENTITY_TYPE_ID"],
                },
            )
            for row in as_list(data.get("result")):
                if str(row.get("ENTITY_TYPE_ID")) == ENTITY_TYPE_COMPANY:
                    return Match(
                        "exact", "company", str(row["ENTITY_ID"]), "", False,
                        "БИН (реквизиты)",
                    )

    # 2. Телефон и email: одним запросом по лидам, контактам и компаниям.
    for kind_name, value in (("PHONE", lead["phone"]), ("EMAIL", lead["email"])):
        if not is_set(value):
            continue
        data = call_bitrix(
            "crm.duplicate.findbycomm",
            {"type": kind_name, "values": split_values(value)},
        )
        found = data.get("result")
        if isinstance(found, dict):
            for key, kind in (("LEAD", "lead"), ("COMPANY", "company"), ("CONTACT", "contact")):
                ids = as_list(found.get(key))
                if ids:
                    return Match(
                        "exact", kind, str(ids[0]), "", False,
                        "телефон" if kind_name == "PHONE" else "email",
                    )

    # 3. Сайт (если фильтр по сайту недоступен на портале - просто пропускаем).
    if is_set(lead["website"]):
        host = website_host(lead["website"])
        try:
            for kind in ("lead", "company"):
                rows = [
                    r for r in list_entities(kind, {"%WEB": host})
                    if host in {
                        website_host(w.get("VALUE") or "")
                        for w in as_list(r.get("WEB"))
                    }
                ]
                if rows:
                    return make_match("exact", kind, rows, lead, "сайт")
        except BitrixError as exc:
            logger.warning("Website duplicate check skipped: %s", exc)

    # 4. Название организации.
    core = normalize_name(company) if is_set(company) else ""
    if len(core) >= 4:
        partial: Optional[Match] = None
        for kind in ("company", "lead"):
            rows = [
                r for r in list_entities(kind, {"%TITLE": core})
                if core in normalize_name(
                    lead_company_part(r["TITLE"]) if kind == "lead" else r["TITLE"]
                )
            ]
            exact_rows = [
                r for r in rows
                if normalize_name(
                    lead_company_part(r["TITLE"]) if kind == "lead" else r["TITLE"]
                ) == core
            ]
            if exact_rows:
                return make_match("exact", kind, exact_rows, lead, "название")
            if rows and partial is None:
                partial = make_match("partial", kind, rows, lead, "название (частично)")
        return partial

    # 5. Заказчик неизвестен: ищем тот же объект по заголовку лида.
    if is_set(lead["project_name"]):
        project = lead["project_name"].lower()
        rows = [
            r for r in list_entities("lead", {"%TITLE": lead["project_name"]})
            if project in (r.get("TITLE") or "").lower()
        ]
        if rows:
            return make_match("exact", "lead", rows, lead, "название объекта")
    return None


def lead_lines(lead: dict) -> list[tuple[str, str]]:
    display = {**lead, "work_type": WORK_TYPE_LABELS.get(lead["work_type"], NOT_SET)}
    lines = [(label, display[key]) for key, label in COMMENT_FIELDS if is_set(display[key])]
    lines.append(("Дата обнаружения", date.today().isoformat()))
    return lines


def build_comments(lead: dict, review: Optional[Match] = None) -> str:
    """HTML-комментарий для карточки лида: краткое описание и ссылки на источники."""
    parts = []
    if review:
        parts.append(
            "<b>ВНИМАНИЕ: возможный дубль.</b> Совпадение по названию с карточкой "
            f"({review.kind} №{review.entity_id}: {html.escape(review.title)}). "
            "Проверьте вручную."
        )
    for label, value in lead_lines(lead):
        text = html.escape(value)
        if label == "Источник" and is_url(value):
            text = f'<a href="{html.escape(value, quote=True)}">{text}</a>'
        parts.append(f"<b>{label}:</b> {text}")
    return "<br>".join(parts)


def build_timeline_text(lead: dict) -> str:
    """Текст комментария для таймлайна существующей карточки (BB-код)."""
    lines = [f"[b]{TITLE_PREFIX} Новая информация, найденная ИИ-поиском[/b]"]
    lines += [f"[b]{label}:[/b] {value}" for label, value in lead_lines(lead)]
    return "\n".join(lines)


def get_source_id() -> str:
    """Код источника лида «ИИ-поиск»: находит в CRM или создаёт."""
    global _source_id_cache
    if BITRIX_SOURCE_ID:
        return BITRIX_SOURCE_ID
    if _source_id_cache:
        return _source_id_cache

    try:
        data = call_bitrix("crm.status.list", {"filter": {"ENTITY_ID": "SOURCE"}})
        for item in as_list(data.get("result")):
            if item.get("NAME") == SOURCE_NAME:
                _source_id_cache = item["STATUS_ID"]
                return _source_id_cache
        call_bitrix(
            "crm.status.add",
            {
                "fields": {
                    "ENTITY_ID": "SOURCE",
                    "STATUS_ID": SOURCE_CODE,
                    "NAME": SOURCE_NAME,
                    "SORT": 1000,
                }
            },
        )
        _source_id_cache = SOURCE_CODE
    except BitrixError as exc:
        logger.warning(
            "Could not find or create lead source '%s' (%s). Using 'OTHER'; "
            "the source is still written to SOURCE_DESCRIPTION.",
            SOURCE_NAME,
            exc,
        )
        _source_id_cache = "OTHER"
    return _source_id_cache


def build_lead_fields(lead: dict, review: Optional[Match] = None) -> dict:
    company = lead["company_name"] if is_set(lead["company_name"]) else "Неизвестный заказчик"
    project = lead["project_name"] if is_set(lead["project_name"]) else "Объект СМР"
    prefix = TITLE_PREFIX + (REVIEW_PREFIX if review else "")

    source_description = SOURCE_NAME
    if is_url(lead["source_link"]):
        source_description += f": {lead['source_link']}"

    fields: dict[str, Any] = {
        "TITLE": f"{prefix} {company} — {project}",
        "STATUS_ID": BITRIX_LEAD_STATUS,
        "SOURCE_ID": get_source_id(),
        "SOURCE_DESCRIPTION": source_description,
        "OPENED": "Y",
        "COMPANY_TITLE": company,
        "COMMENTS": build_comments(lead, review),
    }
    if is_set(lead["contact_person"]):
        fields["NAME"] = lead["contact_person"]
    if is_set(lead["city"]):
        fields["ADDRESS_CITY"] = lead["city"]
    if is_set(lead["region"]):
        fields["ADDRESS_REGION"] = lead["region"]
    if is_set(lead["phone"]):
        fields["PHONE"] = [
            {"VALUE": phone, "VALUE_TYPE": "WORK"} for phone in split_values(lead["phone"])
        ]
    if is_set(lead["email"]):
        fields["EMAIL"] = [
            {"VALUE": mail, "VALUE_TYPE": "WORK"} for mail in split_values(lead["email"])
        ]
    if is_set(lead["website"]):
        fields["WEB"] = [{"VALUE": lead["website"], "VALUE_TYPE": "WORK"}]
    if BITRIX_BIN_FIELD and is_set(lead["bin"]):
        fields[BITRIX_BIN_FIELD] = lead["bin"]
    if BITRIX_ASSIGNED_BY_ID:
        fields["ASSIGNED_BY_ID"] = int(BITRIX_ASSIGNED_BY_ID)
    return fields


def create_lead(lead: dict, review: Optional[Match], dry_run: bool) -> None:
    if dry_run:
        title = f"{lead['company_name']} — {lead['project_name']}"
        logger.info("[DRY-RUN] Would create lead%s: %s", " (проверить дубль)" if review else "", title)
        return
    data = call_bitrix("crm.lead.add", {"fields": build_lead_fields(lead, review)})
    if "result" not in data:
        raise BitrixError("crm.lead.add: в ответе нет ID созданного лида")
    logger.info("Lead created in Bitrix24, ID: %s", data["result"])


def add_info_to_existing(match: Match, lead: dict, dry_run: bool) -> None:
    """Добавляет новую информацию в таймлайн карточки, ничего не удаляя и не меняя.

    Работает для лидов на любой стадии, а также для компаний и контактов.
    """
    if dry_run:
        logger.info("[DRY-RUN] Would add info to %s #%s", match.kind, match.entity_id)
        return
    call_bitrix(
        "crm.timeline.comment.add",
        {
            "fields": {
                "ENTITY_ID": int(match.entity_id),
                "ENTITY_TYPE": match.kind,
                "COMMENT": build_timeline_text(lead),
            }
        },
    )
    logger.info("New info added to %s #%s", match.kind, match.entity_id)


def handle_lead(lead: dict, mode: str, dry_run: bool, stats: RunStats) -> str:
    """Сценарии ТЗ: новая компания / дубль / новая информация / частичное совпадение."""
    label = f"{lead['company_name']} — {lead['project_name']}"
    match = find_existing(lead)
    if match:
        logger.info(
            "Match for '%s': %s, причина: %s, %s #%s «%s», тот же объект: %s",
            label, match.level, match.reason, match.kind, match.entity_id,
            match.title, match.same_project,
        )

    if match is None:
        create_lead(lead, None, dry_run)
        stats.created += 1
        return "created"

    if match.level == "partial":
        logger.info("Partial match for '%s': %s #%s", label, match.kind, match.entity_id)
        create_lead(lead, match, dry_run)
        stats.created += 1
        stats.flagged += 1
        return "review"

    stats.duplicates += 1
    if match.same_project or mode == "skip":
        logger.info(
            "Duplicate excluded: %s (%s, %s #%s)",
            label, match.reason, match.kind, match.entity_id,
        )
        return "excluded"

    add_info_to_existing(match, lead, dry_run)
    stats.updated += 1
    return "updated"


def push_leads(leads: list[dict], stats: RunStats, dry_run: bool, mode: str) -> None:
    state = load_state()
    failed = []
    prefix = "ТЕСТ: " if dry_run else ""

    for lead in leads:
        fingerprint = lead_fingerprint(lead)
        label = f"{lead['company_name']} — {lead['project_name']}"
        if fingerprint in state:
            stats.duplicates += 1
            lead["crm_status"] = prefix + CRM_STATUS_LABELS["known"]
            logger.info("Skipped, already processed on %s: %s", state[fingerprint], label)
            continue
        try:
            outcome = handle_lead(lead, mode, dry_run, stats)
        except BitrixError as exc:
            stats.add_error(f"Bitrix24, «{label}»: {exc}")
            lead["crm_status"] = CRM_STATUS_LABELS["error"]
            failed.append(lead)
            continue
        lead["crm_status"] = prefix + CRM_STATUS_LABELS[outcome]
        if not dry_run:
            state[fingerprint] = date.today().isoformat()
            save_state(state)

    if failed:
        save_failed(failed)


# --------------------------------------------------------------------------
# Запуск
# --------------------------------------------------------------------------
def process_report(
    raw_report: dict,
    source_text: Optional[str],
    stats: RunStats,
    args: argparse.Namespace,
) -> tuple[dict, Path]:
    """Возвращает итоговый отчёт и путь к Excel-файлу."""
    stats.found = len(as_list(raw_report.get("leads")))
    report = normalize_report(raw_report)
    if source_text:
        for lead in report["leads"]:
            verify_lead(lead, source_text, stats)
    report["leads"] = dedupe_batch(report["leads"], stats)
    stats.matched = len(report["leads"])
    for lead in report["leads"]:
        lead["crm_status"] = CRM_NOT_SENT
    logger.info("Found: %d, matched: %d", stats.found, stats.matched)

    retry = bool(args.from_json)
    excel_path = save_outputs(report, retry=retry)  # сохраняем до CRM на случай сбоя

    if args.skip_bitrix:
        logger.info("Bitrix24 push skipped")
        return report, excel_path

    push_leads(report["leads"], stats, args.dry_run, args.duplicates)
    excel_path = save_outputs(report, retry=retry)  # теперь со статусами лидов в CRM
    return report, excel_path


def load_manual_text(path: Path) -> Optional[str]:
    """Читает файл, который вы подготовили вручную (news.txt / news.json)."""
    if not path.exists():
        return None
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        logger.warning("%s is empty", path.name)
        return None
    age_days = (date.today() - date.fromtimestamp(path.stat().st_mtime)).days
    if age_days >= 1:
        logger.warning(
            "%s was last changed %d day(s) ago - make sure it holds today's output",
            path.name,
            age_days,
        )
    return text


def get_source_text(args: argparse.Namespace, stats: RunStats) -> Optional[str]:
    """Текст с найденными проектами: из Gemini API или, если API недоступен, из news.txt."""
    if not args.manual:
        try:
            text = search_projects_text(render_search_prompt())
        except GeminiAccessError as exc:
            logger.warning(
                "Gemini API недоступен (%s). Переключаюсь на ручной режим: %s",
                exc,
                NEWS_TEXT_FILE.name,
            )
        else:
            if text is None:
                stats.add_error("Поиск не вернул результат (см. лог)")
                return None
            save_search_text(text)
            return text

    text = load_manual_text(NEWS_TEXT_FILE)
    if text is None:
        stats.add_error(
            f"Нет доступа к Gemini API, а файл {NEWS_TEXT_FILE.name} пуст или не найден"
        )
    else:
        logger.info("Using manual search text from %s", NEWS_TEXT_FILE.name)
    return text


def get_raw_report(
    source_text: str, system_prompt: str, stats: RunStats
) -> Optional[dict]:
    """JSON-отчёт: из Gemini API или, если API недоступен, из news.json."""
    try:
        report = extract_report_via_gemini(source_text, system_prompt)
    except GeminiAccessError as exc:
        logger.warning(
            "Gemini API недоступен для извлечения (%s). Ищу готовый %s",
            exc,
            NEWS_JSON_FILE.name,
        )
        text = load_manual_text(NEWS_JSON_FILE)
        if text is None:
            stats.add_error(
                "Извлечение через API недоступно (квота/оплата). Получите JSON в "
                f"Google AI Studio с промптом prompt.txt и положите в {NEWS_JSON_FILE.name}"
            )
            return None
        report = parse_json_object(text)
        if report is None:
            stats.add_error(f"{NEWS_JSON_FILE.name} не содержит корректный JSON-объект")
        return report

    if report is None:
        stats.add_error("Не удалось получить JSON-отчёт от модели (см. лог)")
    return report


def run_once(args: argparse.Namespace) -> RunStats:
    """Один полный запуск: поиск, извлечение, проверка, дедубликация, CRM, отчёт."""
    stats = RunStats()
    report: Optional[dict] = None
    excel_path: Optional[Path] = None
    logger.info("Starting pipeline")
    try:
        source_text: Optional[str] = None
        if args.from_json:
            raw_report = json.loads(load_text(args.from_json))
            if not isinstance(raw_report, dict):
                raise ValueError("JSON file must contain an object")
        else:
            system_prompt = load_text(PROMPT_FILE)
            if args.input:
                source_text = load_text(args.input)
            else:
                source_text = get_source_text(args, stats)
                if source_text is None:
                    return stats
            raw_report = get_raw_report(source_text, system_prompt, stats)
            if raw_report is None:
                return stats
        report, excel_path = process_report(raw_report, source_text, stats, args)
    except Exception as exc:  # чтобы ночной запуск не падал целиком
        logger.exception("Pipeline crashed")
        stats.add_error(f"Неожиданная ошибка: {exc}")
    finally:
        save_run_report(stats, not args.skip_bitrix, args.dry_run)
        try:
            notify_telegram(stats, report, excel_path, args)
        except Exception as exc:  # уведомление не должно ронять выгрузку
            logger.warning("Telegram notification failed: %s", type(exc).__name__)
    return stats


def run_daily(at: str, args: argparse.Namespace) -> None:
    """Запуск раз в сутки в указанное время (локальное время сервера)."""
    hour, minute = (int(part) for part in at.split(":"))
    logger.info("Daily mode: next runs at %s every day", at)
    while True:
        now = datetime.now()
        target = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if target <= now:
            target += timedelta(days=1)
        logger.info("Next run: %s", target.strftime("%Y-%m-%d %H:%M"))
        time.sleep((target - now).total_seconds())
        run_once(args)


def schedule_value(value: str) -> str:
    if not re.fullmatch(r"([01]\d|2[0-3]):[0-5]\d", value):
        raise argparse.ArgumentTypeError("use HH:MM, for example 09:00")
    return value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Find construction leads, save JSON and Excel reports, "
            "and push new leads to Bitrix24."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        help="UTF-8 text file with source material. If omitted, the script searches the web.",
    )
    parser.add_argument(
        "--from-json",
        type=Path,
        help="Re-process a saved JSON (for example output/failed_<date>.json) "
        "without calling Gemini. Use it to retry after a technical failure.",
    )
    parser.add_argument(
        "--skip-bitrix",
        action="store_true",
        help="Only save the files, do not touch Bitrix24.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Test run: check duplicates in Bitrix24 (read-only) and log what "
        "would be created, but change nothing in the CRM.",
    )
    parser.add_argument(
        "--duplicates",
        choices=("update", "skip"),
        default="update",
        help="update: add new info about a known company to its existing card "
        "(timeline comment, nothing is deleted); skip: just exclude the duplicate.",
    )
    parser.add_argument(
        "--manual",
        action="store_true",
        help="Do not call the Gemini search; take the search text from news.txt.",
    )
    parser.add_argument(
        "--print-search-prompt",
        action="store_true",
        help="Print today's search prompt (to paste into Google AI Studio) and exit.",
    )
    parser.add_argument(
        "--test-telegram",
        action="store_true",
        help="Send a test message to the Telegram chat and exit.",
    )
    parser.add_argument(
        "--schedule",
        type=schedule_value,
        metavar="HH:MM",
        help="Run automatically once a day at this time and keep running.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.print_search_prompt:
        print(render_search_prompt())
        return
    if args.test_telegram:
        if not telegram_enabled():
            sys.exit("TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID must be set in .env")
        ok = send_telegram_message("✅ CLC: бот подключён, уведомления будут приходить сюда.")
        sys.exit(0 if ok else 1)
    if not args.skip_bitrix and not BITRIX_WEBHOOK_URL:
        sys.exit("BITRIX_WEBHOOK_URL is missing in .env (or use --skip-bitrix)")
    if args.schedule and (args.from_json or args.input):
        sys.exit("--schedule cannot be combined with --input or --from-json")

    if args.schedule:
        run_daily(args.schedule, args)
        return

    stats = run_once(args)
    sys.exit(1 if stats.errors else 0)


if __name__ == "__main__":
    main()