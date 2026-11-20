import argparse
import hashlib
import html
import json
import logging
import os
import re
import sys
import time
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
SEARCH_PROMPT_FILE = BASE_DIR / "Search prompt.txt"
CRITERIA_FILE = BASE_DIR / "criteria.json"
OUTPUT_DIR = BASE_DIR / "output"
STATE_FILE = OUTPUT_DIR / "state.json"

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

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").rstrip("/")
SEARCH_MODEL = os.getenv("OPENAI_SEARCH_MODEL", "gpt-4o-mini")
EXTRACT_MODEL = os.getenv("OPENAI_EXTRACT_MODEL", "gpt-4o-mini")

BITRIX_WEBHOOK_URL = os.getenv("BITRIX_WEBHOOK_URL")
BITRIX_LEAD_STATUS = os.getenv("BITRIX_LEAD_STATUS", "NEW")
BITRIX_SOURCE_ID = os.getenv("BITRIX_SOURCE_ID")
BITRIX_BIN_FIELD = os.getenv("BITRIX_BIN_FIELD")
BITRIX_REQUISITE_BIN_FIELD = os.getenv("BITRIX_REQUISITE_BIN_FIELD")
BITRIX_ASSIGNED_BY_ID = os.getenv("BITRIX_ASSIGNED_BY_ID")
BITRIX_PAUSE = float(os.getenv("BITRIX_PAUSE", "0.5"))

TELEGRAM_BOT_TOKEN = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip().strip("\"'")
TELEGRAM_CHAT_ID = (os.getenv("TELEGRAM_CHAT_ID") or "").strip().strip("\"'")
TELEGRAM_API = "https://api.telegram.org/bot{token}/{method}"
TELEGRAM_TIMEOUT = 60
TELEGRAM_MAX_LEADS_IN_MESSAGE = 15

OPENAI_TIMEOUT = int(os.getenv("OPENAI_TIMEOUT", "120"))
OPENAI_MAX_TOKENS = int(os.getenv("OPENAI_MAX_TOKENS", "16384"))
BITRIX_TIMEOUT = 15
BITRIX_RETRIES = 3
BIN_LENGTH = 12
NOT_SET = "Не установлено"

SOURCE_NAME = "ИИ-поиск"
SOURCE_CODE = "AI_SEARCH"
TITLE_PREFIX = "[AI-CLC]"
REVIEW_PREFIX = "[ПРОВЕРИТЬ ДУБЛЬ]"

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
    ("offer_deadline", "Срок предложения", 22),
    ("budget", "Бюджет", 20),
    ("project_scale", "Масштаб проекта", 24),
    ("clc_demand", "Потребность в продукции CLC", 40),
    ("contact_person", "Контактное лицо", 24),
    ("phone", "Телефон", 20),
    ("email", "Email", 26),
    ("source_link", "Источник", 40),
    ("publication_date", "Дата публикации", 18),
    ("recommendation", "Рекомендуемые действия", 45),
    ("expected_result", "Ожидаемый результат", 40),
    ("risk_check", "Риски / что проверить", 45),
    ("priority", "Приоритет", 12),
    ("crm_status", "Статус в CRM", 30),
]
LEAD_FIELDS = [key for key, _, _ in LEAD_COLUMNS]
ANALYTICS_FIELDS = ("project_name", "fact", "impact", "action")
TECH_NEWS_FIELDS = (
    "date",
    "initiative",
    "verified_fact",
    "scale",
    "clc_connection",
    "limitation",
    "action",
    "source_link",
)

RADAR_COLUMNS = [
    ("number", "№", 6),
    ("publication_date", "Дата", 16),
    ("signal_type", "Тип", 24),
    ("project_name", "Сигнал", 40),
    ("location", "Локация", 25),
    ("verified_fact", "Проверенный факт", 49),
    ("clc_demand", "Возможность CLC", 52),
    ("rating", "Приоритет", 14),
    ("source_link", "Источник", 44),
]
PROJECT_COLUMNS = [
    ("number", "№", 6),
    ("project_label", "Проект", 38),
    ("rating", "Оценка", 13),
    ("publication_date", "Дата", 15),
    ("location", "Локация", 25),
    ("stage", "Стадия", 34),
    ("scale", "Бюджет / масштаб", 30),
    ("company_name", "Заказчик / куратор", 34),
    ("contractor", "Подрядчик / проектировщик", 31),
    ("clc_demand", "Позиции CLC", 54),
    ("contact", "Кому обращаться", 35),
    ("phone", "Телефон", 22),
    ("email", "Email", 28),
    ("offer_deadline", "Срок предложения", 20),
    ("recommendation", "Первый шаг", 45),
    ("risk_check", "Риски / проверка", 48),
    ("source_link", "Источник", 44),
]
EDUCATION_COLUMNS = [
    ("number", "№", 6),
    ("project_name", "Объект / программа", 36),
    ("object_type", "Тип", 24),
    ("region", "Регион", 22),
    ("scale", "Масштаб", 24),
    ("stage", "Стадия", 33),
    ("current_status", "Актуальный статус", 38),
    ("company_name", "Заказчик / канал", 34),
    ("contractor", "Подрядчик / ГИП", 31),
    ("clc_demand", "Позиции CLC", 52),
    ("procurement_status", "Закупочный статус", 30),
    ("recommendation", "Действие", 47),
    ("rating", "Оценка", 14),
    ("repair_control", "Контроль ремонта", 30),
    ("risk_check", "Ограничение", 45),
    ("source_link", "Источник", 44),
]
TODAY_COLUMNS = [
    ("number", "Приоритет", 10),
    ("project_label", "Лид", 42),
    ("responsible", "Ответственный", 31),
    ("recommendation", "Действие", 59),
    ("contact", "Контакт", 43),
    ("offer_deadline", "Дедлайн", 22),
    ("expected_result", "Ожидаемый результат", 53),
    ("task_status", "Статус", 18),
]
TECH_COLUMNS = [
    ("date", "Дата", 15),
    ("initiative", "Инициатива", 38),
    ("verified_fact", "Проверенный факт", 56),
    ("scale", "Масштаб", 34),
    ("clc_connection", "Связь с CLC", 52),
    ("limitation", "Ограничение", 48),
    ("action", "Действие", 52),
    ("source_link", "Источник", 44),
]

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
    ("offer_deadline", "Срок предложения"),
    ("budget", "Бюджет"),
    ("project_scale", "Масштаб проекта"),
    ("clc_demand", "Потребность в продукции CLC"),
    ("contact_person", "Контактное лицо"),
    ("phone", "Телефон"),
    ("email", "Email"),
    ("publication_date", "Дата публикации"),
    ("priority", "Приоритет"),
    ("recommendation", "Рекомендация CLC"),
    ("expected_result", "Ожидаемый результат"),
    ("risk_check", "Риски / что проверить"),
    ("source_link", "Источник"),
]

LEGAL_FORMS = {
    "тоо", "ао", "кгу", "ргу", "гу", "кгп", "ргп", "ип", "оао", "зао", "ооо",
    "llp", "jsc", "тов",
}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_set(value: Optional[str]) -> bool:
    return bool(value) and value != NOT_SET


def as_text(value: Any) -> str:
    if value is None:
        return NOT_SET
    text = str(value).strip()
    return text or NOT_SET


def as_list(value: Any) -> list:
    return value if isinstance(value, list) else []


def is_url(value: str) -> bool:
    return value.startswith(("http://", "https://")) and len(value) < 2000


def digits(value: str) -> str:
    return re.sub(r"\D", "", value)


def split_values(value: str) -> list[str]:
    return [part.strip() for part in re.split(r"[;,\n]+", value) if part.strip()]


def website_host(value: str) -> str:
    url = value if "://" in value else f"https://{value}"
    host = urlparse(url).netloc.lower()
    return host[4:] if host.startswith("www.") else host


def normalize_name(value: str) -> str:
    text = value.lower().replace("ё", "е")
    text = re.sub(r"[«»\"'“”„().,;:–—-]", " ", text)
    return " ".join(word for word in text.split() if word not in LEGAL_FORMS)


def lead_fingerprint(lead: dict) -> str:
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


@dataclass
class RunStats:
    found: int = 0
    matched: int = 0
    created: int = 0
    flagged: int = 0
    duplicates: int = 0
    updated: int = 0
    unverified_removed: int = 0
    errors: int = 0
    error_details: list = field(default_factory=list)

    def add_error(self, message: str) -> None:
        self.errors += 1
        self.error_details.append(message)
        logger.error(message)


def load_state() -> dict:
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


class OpenAIAccessError(Exception):
    """Нет доступа к OpenAI API: квота, оплата, права или не задан ключ."""


def is_access_problem(status: int, body: str) -> bool:
    if status in (402, 403, 429):
        return True
    lowered = body.lower()
    return status == 400 and ("failed_precondition" in lowered or "billing" in lowered)


def call_openai(
    model: str,
    prompt: str,
    system_prompt: Optional[str] = None,
    use_search: bool = False,
    json_mode: bool = False,
) -> Optional[str]:
    if not OPENAI_API_KEY:
        raise OpenAIAccessError("OPENAI_API_KEY не задан")
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }

    if use_search:
        url = f"{OPENAI_BASE_URL}/responses"
        payload: dict[str, Any] = {
            "model": model,
            "input": prompt,
            "tools": [
                {
                    "type": "web_search",
                    "search_context_size": "high",
                    "user_location": {"type": "approximate", "country": "KZ"},
                }
            ],
            "max_output_tokens": OPENAI_MAX_TOKENS,
        }
        if system_prompt:
            payload["instructions"] = system_prompt
        if json_mode:
            payload["text"] = {"format": {"type": "json_object"}}
    else:
        url = f"{OPENAI_BASE_URL}/chat/completions"
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        payload = {
            "model": model,
            "messages": messages,
            "max_tokens": OPENAI_MAX_TOKENS,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}

    try:
        response = requests.post(
            url, headers=headers, json=payload, timeout=OPENAI_TIMEOUT
        )
        response.raise_for_status()
        data = response.json()
    except requests.HTTPError as exc:
        status = exc.response.status_code
        if is_access_problem(status, exc.response.text):
            raise OpenAIAccessError(
                f"HTTP {status} для модели {model}"
            ) from exc
        logger.error(
            "OpenAI returned HTTP %s for model '%s': %s",
            status,
            model,
            exc.response.text[:300],
        )
        return None
    except requests.RequestException as exc:
        logger.error("OpenAI request failed for model '%s': %s", model, exc)
        return None
    except ValueError as exc:
        logger.error("Unexpected OpenAI response format: %s", exc)
        return None

    if use_search:
        output_text = data.get("output_text")
        if output_text:
            return output_text.strip()
        texts = []
        for item in data.get("output", []):
            for content in item.get("content", []):
                if content.get("type") in ("output_text", "text"):
                    texts.append(content.get("text", ""))
        return "".join(texts).strip() or None

    try:
        return data["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError, TypeError) as exc:
        logger.error("Unexpected OpenAI chat response format: %s", exc)
        return None


def load_criteria() -> dict:
    try:
        data = json.loads(load_text(CRITERIA_FILE))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def render_search_prompt() -> str:
    criteria = load_criteria()
    days = int(criteria.get("days", 14))
    max_projects = int(criteria.get("max_projects", 10))
    today = date.today()
    since = today - timedelta(days=days)
    template = load_text(SEARCH_PROMPT_FILE)

    def join(key: str) -> str:
        return "; ".join(criteria.get(key, []))

    common = {
        "<<TODAY>>": today.isoformat(),
        "<<SINCE>>": since.isoformat(),
        "<<DAYS>>": str(days),
        "<<MAX>>": str(max_projects),
        "<<OBJECT_TYPES>>": join("object_types"),
        "<<WORK_TYPES>>": join("work_types"),
        "<<KEYWORDS>>": join("keywords"),
        "<<EXCLUDE>>": join("exclude_keywords"),
        "<<SOURCES>>": join("sources"),
    }

    regions = "; ".join(criteria.get("regions") or ["весь Казахстан"])
    focus = (
        "Один сводный поиск по всему Казахстану, охвати все регионы: "
        f"{regions}. В первую очередь ищи частных застройщиков, девелоперов, "
        "владельцев коммерческой недвижимости и другие частные компании. "
        "Не разбивай поиск на отдельные задания по регионам."
    )
    prompt = template.replace("<<FOCUS>>", focus)
    for token, value in common.items():
        prompt = prompt.replace(token, value)
    return prompt


def search_projects_text(search_prompt: str) -> Optional[str]:
    logger.info("Searching the web with OpenAI model '%s'", SEARCH_MODEL)
    text = call_openai(SEARCH_MODEL, search_prompt, use_search=True)
    if text is None:
        return None
    if not text:
        logger.error("OpenAI returned an empty answer for the search step")
        return None
    return text


def parse_json_object(content: str) -> Optional[dict]:
    start = content.find("{")
    end = content.rfind("}")
    if start == -1 or end <= start:
        return None
    try:
        parsed = json.loads(content[start : end + 1])
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def extract_report_via_openai(raw_text: str, system_prompt: str) -> Optional[dict]:
    prompt = (
        f"Today is {date.today().isoformat()}. "
        "Extract construction leads for CLC Company from the text below and "
        "return the report as a JSON object that follows the schema from the "
        f"system prompt.\n\n{raw_text}"
    )
    content = call_openai(
        EXTRACT_MODEL, prompt, system_prompt=system_prompt, json_mode=True
    )
    if content is None:
        return None
    report = parse_json_object(content)
    if report is None:
        logger.error("Model answer does not contain a valid JSON object")
    return report


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
    leads = [
        normalize_lead(item)
        for item in as_list(raw.get("leads"))
        if isinstance(item, dict)
    ]
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
    raw_technology_news = raw_summary.get("technology_news")
    if isinstance(raw_technology_news, dict):
        technology_news = {
            key: as_text(raw_technology_news.get(key))
            for key in TECH_NEWS_FIELDS
        }
    else:
        technology_news = {key: NOT_SET for key in TECH_NEWS_FIELDS}
    tech_news = as_text(raw_summary.get("tech_news"))
    if not is_set(tech_news) and is_set(technology_news["verified_fact"]):
        tech_news = technology_news["verified_fact"]

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
            "tech_news": tech_news,
            "technology_news": technology_news,
        },
    }


def dedupe_batch(leads: list[dict], stats: RunStats) -> list[dict]:
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


def put(
    ws: Worksheet,
    row: int,
    column: int,
    value: Any,
    font: Font = BODY_FONT,
    formula: bool = False,
):
    cell = ws.cell(row=row, column=column, value=value)
    if isinstance(value, str) and value.startswith("=") and not formula:
        cell.data_type = "s"
    cell.font = font
    cell.alignment = CELL_ALIGNMENT
    return cell


def fill_table_sheet(
    ws: Worksheet, columns: list, rows: list[dict], header_row: int = 1
) -> None:
    for index, (_, header, width) in enumerate(columns, start=1):
        cell = put(ws, header_row, index, header, font=HEADER_FONT)
        cell.fill = HEADER_FILL
        ws.column_dimensions[get_column_letter(index)].width = width

    for row_number, row in enumerate(rows, start=1):
        row_index = header_row + row_number
        for col_index, (key, _, _) in enumerate(columns, start=1):
            value = row_number if key == "number" else row.get(key, NOT_SET)
            cell = put(ws, row_index, col_index, value)
            if key == "source_link" and is_url(value):
                cell.hyperlink = value
                cell.font = LINK_FONT
            elif key in ("priority", "rating"):
                priority = value
                if value in ("Горячий", "Теплый", "Тёплый", "Наблюдать"):
                    priority = {
                        "Горячий": "A",
                        "Теплый": "B",
                        "Тёплый": "B",
                        "Наблюдать": "C",
                    }[value]
                cell.fill = PatternFill(
                    "solid", start_color=PRIORITY_COLORS.get(priority, "E7E6E6")
                )

    last_column = get_column_letter(len(columns))
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = (
        f"A{header_row}:{last_column}{max(header_row + len(rows), header_row + 1)}"
    )


def to_display_rows(leads: list[dict]) -> list[dict]:
    return [
        {**lead, "work_type": WORK_TYPE_LABELS.get(lead["work_type"], NOT_SET)}
        for lead in leads
    ]


def rating_label(priority: str) -> str:
    return {"A": "Горячий", "B": "Теплый", "C": "Наблюдать"}.get(
        priority, NOT_SET
    )


def report_location(lead: dict) -> str:
    parts = []
    for key in ("city", "region"):
        value = lead[key]
        if is_set(value) and value not in parts:
            parts.append(value)
    return ", ".join(parts) or NOT_SET


def verified_fact(lead: dict) -> str:
    facts = [
        f"Заказчик: {lead['company_name']}"
        if is_set(lead["company_name"])
        else "",
        f"Генподрядчик: {lead['general_contractor']}"
        if is_set(lead["general_contractor"])
        else "",
        f"Стадия: {lead['stage']}" if is_set(lead["stage"]) else "",
        f"Сроки: {lead['timeline']}" if is_set(lead["timeline"]) else "",
        f"Бюджет: {lead['budget']}" if is_set(lead["budget"]) else "",
        f"Масштаб: {lead['project_scale']}"
        if is_set(lead["project_scale"])
        else "",
    ]
    return "; ".join(fact for fact in facts if fact) or NOT_SET


def report_scale(lead: dict) -> str:
    parts = [
        value
        for value in (lead["budget"], lead["project_scale"])
        if is_set(value)
    ]
    return "; ".join(parts) or NOT_SET


def project_contractor(lead: dict) -> str:
    contractors = [
        value
        for value in (
            lead["general_contractor"],
            lead["finishing_subcontractor"],
        )
        if is_set(value)
    ]
    return "; ".join(contractors) or NOT_SET


def project_contact(lead: dict) -> str:
    for key in ("contact_person", "general_contractor", "company_name"):
        if is_set(lead[key]):
            return lead[key]
    return NOT_SET


def project_risks(lead: dict) -> str:
    if is_set(lead["risk_check"]):
        return lead["risk_check"]
    missing = [
        label
        for key, label in (
            ("general_contractor", "генподрядчик"),
            ("contact_person", "контактное лицо"),
            ("phone", "телефон"),
            ("offer_deadline", "срок предложения"),
            ("budget", "бюджет"),
        )
        if not is_set(lead[key])
    ]
    if missing:
        return "Не установлено в источниках: " + ", ".join(missing)
    return NOT_SET


def report_date_label(value: str) -> str:
    months = (
        "января",
        "февраля",
        "марта",
        "апреля",
        "мая",
        "июня",
        "июля",
        "августа",
        "сентября",
        "октября",
        "ноября",
        "декабря",
    )
    try:
        parsed = date.fromisoformat(value)
    except ValueError:
        return value
    return f"{parsed.day} {months[parsed.month - 1]} {parsed.year} года"


def radar_summary(report: dict) -> str:
    leads = report["leads"]
    if not leads:
        return "За выбранный период новых подтвержденных лидов не найдено."

    priority_names = [
        lead["project_name"] for lead in leads if lead["priority"] == "A"
    ][:3]
    actions = as_list(report["summary"].get("priority_actions"))[:2]
    parts = [f"Найдено частных проектов: {len(leads)}."]
    if priority_names:
        parts.append("Приоритетно: " + "; ".join(priority_names) + ".")
    if actions:
        action_text = "; ".join(
            action.rstrip(" .;") for action in actions if isinstance(action, str)
        )
        if action_text:
            parts.append("Первые действия: " + action_text + ".")
    return " ".join(parts)


def write_report_sheet(
    ws: Worksheet,
    title: str,
    columns: list,
    rows: list[dict],
    subtitle: Optional[str] = None,
) -> None:
    last_column = get_column_letter(len(columns))
    ws.merge_cells(f"A1:{last_column}1")
    put(ws, 1, 1, title, font=TITLE_FONT)
    header_row = 3
    if subtitle is not None:
        ws.merge_cells(f"A2:{last_column}2")
        put(ws, 2, 1, subtitle)
        header_row = 4
    fill_table_sheet(ws, columns, rows, header_row=header_row)


def build_excel(report: dict, path: Path) -> None:
    leads = report["leads"]
    workbook = Workbook()

    lead_by_name = {lead["project_name"]: lead for lead in leads}
    summary = report["summary"]
    top_names = [
        name
        for name in as_list(summary.get("top_projects"))
        if name in lead_by_name
    ][:4]
    if not top_names:
        top_names = [lead["project_name"] for lead in leads[:4]]

    radar_rows = []
    for lead in leads:
        work_type = lead["work_type"]
        if work_type in REPAIR_SHEET_TYPES:
            signal_type = "РЕМОНТ / РЕКОНСТРУКЦИЯ"
        elif is_set(lead["general_contractor"]):
            signal_type = "ГЕНПОДРЯДЧИК ВЫБРАН"
        else:
            signal_type = "ЧАСТНЫЙ ПРОЕКТ"
        radar_rows.append(
            {
                "publication_date": lead["publication_date"],
                "signal_type": signal_type,
                "project_name": lead["project_name"],
                "location": report_location(lead),
                "verified_fact": verified_fact(lead),
                "clc_demand": lead["clc_demand"],
                "rating": rating_label(lead["priority"]),
                "source_link": lead["source_link"],
            }
        )

    radar_sheet = workbook.active
    radar_sheet.title = "Радар"
    write_report_sheet(
        radar_sheet,
        f"CLC Company — деловая сводка, {report_date_label(report['report_date'])}",
        RADAR_COLUMNS,
        radar_rows,
        subtitle=radar_summary(report),
    )

    project_rows = []
    for name in top_names:
        lead = lead_by_name[name]
        project_rows.append(
            {
                "project_label": lead["project_name"],
                "rating": rating_label(lead["priority"]),
                "publication_date": lead["publication_date"],
                "location": report_location(lead),
                "stage": lead["stage"],
                "scale": report_scale(lead),
                "company_name": lead["company_name"],
                "contractor": project_contractor(lead),
                "clc_demand": lead["clc_demand"],
                "contact": project_contact(lead),
                "phone": lead["phone"],
                "email": lead["email"],
                "offer_deadline": lead["offer_deadline"],
                "recommendation": lead["recommendation"],
                "risk_check": project_risks(lead),
                "source_link": lead["source_link"],
            }
        )
    projects_sheet = workbook.create_sheet("Проекты")
    write_report_sheet(
        projects_sheet,
        "Четыре приоритетных направления продаж CLC",
        PROJECT_COLUMNS,
        project_rows,
    )

    education_terms = (
        "школ",
        "детск",
        "больниц",
        "медиц",
        "клиник",
        "колледж",
        "универ",
    )
    education_leads = [
        lead
        for lead in leads
        if any(term in lead["object_type"].lower() for term in education_terms)
        or lead["work_type"] in REPAIR_SHEET_TYPES
    ]
    education_rows = []
    for lead in education_leads:
        status_parts = [
            lead[key]
            for key in ("stage", "timeline")
            if is_set(lead[key])
        ]
        education_rows.append(
            {
                "project_name": lead["project_name"],
                "object_type": lead["object_type"],
                "region": lead["region"],
                "scale": report_scale(lead),
                "stage": lead["stage"],
                "current_status": "; ".join(status_parts) or NOT_SET,
                "company_name": lead["company_name"],
                "contractor": project_contractor(lead),
                "clc_demand": lead["clc_demand"],
                "procurement_status": (
                    "Генподрядчик подтвержден"
                    if is_set(lead["general_contractor"])
                    else "Генподрядчик не установлен в источниках"
                ),
                "recommendation": lead["recommendation"],
                "rating": rating_label(lead["priority"]),
                "repair_control": NOT_SET,
                "risk_check": project_risks(lead),
                "source_link": lead["source_link"],
            }
        )
    education_sheet = workbook.create_sheet("Образование и ремонты")
    write_report_sheet(
        education_sheet,
        f"Образовательные, медицинские и ремонтные сигналы — {report_date_label(report['report_date'])}",
        EDUCATION_COLUMNS,
        education_rows,
    )

    today_leads = [
        lead
        for lead in leads
        if lead["priority"] == "A" or is_set(lead["offer_deadline"])
    ]
    today_rows = []
    for lead in today_leads:
        contact_details = [
            value
            for value in (lead["contact_person"], lead["phone"], lead["email"])
            if is_set(value)
        ]
        today_rows.append(
            {
                "project_label": lead["project_name"],
                "responsible": "Назначить менеджера региона",
                "recommendation": lead["recommendation"],
                "contact": "; ".join(contact_details) or project_contact(lead),
                "offer_deadline": lead["offer_deadline"],
                "expected_result": lead["expected_result"],
                "task_status": (
                    "Критично"
                    if is_set(lead["offer_deadline"])
                    else {
                        "A": "Высокий",
                        "B": "Плановый",
                        "C": "Наблюдать",
                    }.get(lead["priority"], "Не установлено")
                ),
            }
        )
    today_sheet = workbook.create_sheet("Отработать сегодня")
    write_report_sheet(
        today_sheet,
        "Отработать сегодня — задачи для CRM",
        TODAY_COLUMNS,
        today_rows,
    )

    technology_news = summary.get("technology_news", {})
    if not isinstance(technology_news, dict):
        technology_news = {}
    technology_row = {
        key: as_text(technology_news.get(key)) for key in TECH_NEWS_FIELDS
    }
    if not is_set(technology_row["verified_fact"]):
        technology_row["verified_fact"] = as_text(summary.get("tech_news"))
    if not is_set(technology_row["initiative"]) and is_set(
        technology_row["verified_fact"]
    ):
        technology_row["initiative"] = "Технологическая новость"
    technology_rows = [technology_row]
    technology_title = (
        "Технологии и ИИ — "
        + (
            technology_row["initiative"]
            if is_set(technology_row["initiative"])
            else "новость не установлена"
        )
    )
    technology_sheet = workbook.create_sheet("Технологии и ИИ")
    write_report_sheet(
        technology_sheet,
        technology_title,
        TECH_COLUMNS,
        technology_rows,
    )

    workbook.save(path)


def save_search_text(text: str, name: str = "all") -> None:
    safe = re.sub(r"[^\w-]+", "_", name, flags=re.UNICODE).strip("_")[:40] or "all"
    path = OUTPUT_DIR / f"search_{date.today().isoformat()}_{datetime.now():%H%M%S}_{safe}.txt"
    path.write_text(text, encoding="utf-8")
    logger.info("Search text saved: %s", path)


def save_outputs(report: dict, retry: bool = False) -> Path:
    stamp = report["report_date"]

    if not retry:
        json_path = OUTPUT_DIR / f"leads_{stamp}.json"
        json_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        logger.info("JSON saved: %s", json_path)

    filename = (
        f"CLC_деловая_сводка_retry_{stamp}.xlsx"
        if retry
        else f"CLC_деловая_сводка_{stamp}.xlsx"
    )
    excel_path = OUTPUT_DIR / filename
    build_excel(report, excel_path)
    logger.info("Excel saved: %s", excel_path)
    return excel_path


def save_failed(leads: list[dict]) -> None:
    path = OUTPUT_DIR / f"failed_{date.today().isoformat()}.json"
    path.write_text(
        json.dumps({"leads": leads, "summary": {}}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.warning(
        "%d lead(s) were not delivered. Retry with: --from-json %s", len(leads), path
    )


def telegram_enabled() -> bool:
    return bool(TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID)


def telegram_call(method: str, data: dict, files: Optional[dict] = None) -> bool:
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
        lines = [f" CLC: в CRM Bitrix24 добавлено лидов: {stats.created}"]
    else:
        lines = ["ℹCLC: новых лидов в CRM не добавлено"]
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
        lines.append(f"Ошибок: {stats.errors}")
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
    if args.dry_run or args.skip_bitrix:
        return
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


class BitrixError(Exception):
    """Ошибка вызова REST API Bitrix24."""


_last_bitrix_call = 0.0
_source_id_cache: Optional[str] = None


def call_bitrix(method: str, payload: dict) -> dict[str, Any]:
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
            last_error = type(exc).__name__
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
    level: str
    kind: str
    entity_id: str
    title: str
    same_project: bool
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
    for prefix in (REVIEW_PREFIX, TITLE_PREFIX):
        title = title.replace(prefix, "")
    return title.split(" — ")[0].strip()


def find_existing(lead: dict) -> Optional[Match]:
    company, bin_code = lead["company_name"], lead["bin"]
    if is_set(bin_code):
        for kind in ("lead", "company"):
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
    lines = [f"[b]{TITLE_PREFIX} Новая информация, найденная ИИ-поиском[/b]"]
    lines += [f"[b]{label}:[/b] {value}" for label, value in lead_lines(lead)]
    return "\n".join(lines)


def get_source_id() -> str:
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
    bin_value = lead.get("bin", "Не установлено")

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

    if BITRIX_BIN_FIELD and is_set(bin_value):
        fields[BITRIX_BIN_FIELD] = bin_value

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


def process_report(
    raw_report: dict,
    source_text: Optional[str],
    stats: RunStats,
    args: argparse.Namespace,
) -> tuple[dict, Path]:
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
    excel_path = save_outputs(report, retry=retry)

    if args.skip_bitrix:
        logger.info("Bitrix24 push skipped")
        return report, excel_path

    push_leads(report["leads"], stats, args.dry_run, args.duplicates)
    excel_path = save_outputs(report, retry=retry)
    return report, excel_path


def get_source_text(stats: RunStats) -> Optional[str]:
    try:
        text = search_projects_text(render_search_prompt())
    except OpenAIAccessError as exc:
        stats.add_error(f"OpenAI API недоступен: {exc}")
        return None
    if text is None:
        stats.add_error("Поиск не вернул результат (см. лог)")
        return None
    save_search_text(text)
    return text


def get_raw_report(
    source_text: str, system_prompt: str, stats: RunStats
) -> Optional[dict]:
    try:
        report = extract_report_via_openai(source_text, system_prompt)
    except OpenAIAccessError as exc:
        stats.add_error(f"OpenAI API недоступен для извлечения: {exc}")
        return None

    if report is None:
        stats.add_error("Не удалось получить JSON-отчёт от модели (см. лог)")
    return report


def run_once(args: argparse.Namespace) -> RunStats:
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
                source_text = get_source_text(stats)
                if source_text is None:
                    return stats
            raw_report = get_raw_report(source_text, system_prompt, stats)
            if raw_report is None:
                return stats
        report, excel_path = process_report(raw_report, source_text, stats, args)
    except Exception as exc:
        logger.exception("Pipeline crashed")
        stats.add_error(f"Неожиданная ошибка: {exc}")
    finally:
        save_run_report(stats, not args.skip_bitrix, args.dry_run)
        try:
            notify_telegram(stats, report, excel_path, args)
        except Exception as exc:
            logger.warning("Telegram notification failed: %s", type(exc).__name__)
    return stats


def run_daily(at: str, args: argparse.Namespace) -> None:
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
        "without calling OpenAI. Use it to retry after a technical failure.",
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
        "--print-search-prompt",
        action="store_true",
        help="Print the prompt.txt search prompt and exit.",
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
        ok = send_telegram_message(" CLC: бот подключён, уведомления будут приходить сюда.")
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
