"""
Document analysis for LetterCheck Smart.

The analyzer keeps the old public function name for compatibility with the
FastAPI layer, but internally works in three stages:
1. Read the document in body order.
2. Classify blocks: header/addressee, greeting, body, signature, executor.
3. Run checks over the classified text.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


LLM_BASE_URL = "http://10.66.80.31:8000/v1"
LLM_MODEL = "Qwen3.5-27B"

_llm_client = None


FIO_SURNAME_INITIALS_RE = re.compile(
    r"\b([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)[ \t]+([А-ЯЁ])\.[ \t]*([А-ЯЁ])?\.?"
)
FIO_INITIALS_SURNAME_RE = re.compile(
    r"\b([А-ЯЁ])\.[ \t]*([А-ЯЁ])?\.?[ \t]*([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)"
)
FULL_NAME_RE = re.compile(
    r"\b([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)[ \t]+"
    r"([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)[ \t]+"
    r"([А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?)\b"
)
GREETING_RE = re.compile(r"^\s*(Уважаем(?:ый|ая|ые|ое)\b[^.!?\n]*(?:[!.])?)", re.IGNORECASE)
COLLECTIVE_WORDS = (
    "коллеги",
    "коллег",
    "господа",
    "дамы и господа",
    "товарищи",
    "руководители",
)
BODY_START_RE = re.compile(
    r"\b(направля|просим|сообща|уведомля|представля|рассмотрев|в ответ|в соответствии|согласно|на основании)\b",
    re.IGNORECASE,
)
EXECUTOR_RE = re.compile(
    r"\b(исп\.?|исполнитель|тел\.?|телефон|e-?mail|электронная почта)\b|"
    r"\b8\s*\(\d{3}\)\s*\d{2,3}[-\s]?\d{2}[-\s]?\d{2}\b|"
    r"\(\s*доб\.?\s*\d+\s*\)",
    re.IGNORECASE,
)
SIGNATURE_RE = re.compile(
    r"\b(руководитель|директор|начальник|заместитель|председатель|представитель|судья|прокурор|министр)\b",
    re.IGNORECASE,
)
NORMATIVE_RE = re.compile(
    r"\b("
    r"(?:федеральн(?:ый|ого)\s+закон|закон|постановлени[ея]|приказ|распоряжени[ея]|"
    r"кодекс|гк\s+рф|гпк\s+рф|апк\s+рф|кас\s+рф|коап\s+рф)"
    r"[^\n;]{0,220}"
    r")",
    re.IGNORECASE,
)

MALE_PATRONYMIC_ENDINGS = ("ич", "лы", "улы")
FEMALE_PATRONYMIC_ENDINGS = ("на", "кызы")
COMMON_FEMALE_NAMES = {
    "александра", "алена", "алина", "алла", "анастасия", "анна", "валентина",
    "валерия", "вера", "виктория", "галина", "дарья", "екатерина", "елена",
    "жанна", "зоя", "инна", "ирина", "кристина", "ксения", "лариса", "любовь",
    "людмила", "марина", "мария", "наталья", "нина", "оксана", "ольга",
    "полина", "светлана", "татьяна", "юлия", "яна",
}
COMMON_MALE_NAMES = {
    "александр", "алексей", "андрей", "антон", "артем", "борис", "вадим",
    "валерий", "василий", "виктор", "виталий", "владимир", "геннадий",
    "георгий", "дмитрий", "евгений", "егор", "иван", "игорь", "илья",
    "кирилл", "константин", "максим", "михаил", "николай", "олег", "павел",
    "петр", "роман", "сергей", "станислав", "федор", "юрий",
}

BLOCK_LABELS = {
    "has_header": "Шапка",
    "has_addressee": "Адресат",
    "has_greeting": "Обращение",
    "has_body": "Основной текст",
    "has_signature": "Подпись",
    "has_executor": "Отметка об исполнителе",
    "has_attachment": "Отметка о приложении",
}
BLOCK_ORDER = (
    "has_header",
    "has_addressee",
    "has_greeting",
    "has_body",
    "has_signature",
    "has_executor",
    "has_attachment",
)
DOCUMENT_TYPE_LABELS = {
    "outgoing_letter": "Исходящее письмо",
    "outgoing_without_greeting": "Исходящее письмо без обращения",
    "court_letter": "Письмо в суд",
}
TEXT_RULES = (
    {
        "pattern": re.compile(r"\bв течении\b", re.IGNORECASE),
        "corrected": "в течение",
        "type": "spelling",
        "description": "В значении периода времени используется форма «в течение».",
    },
    {
        "pattern": re.compile(r"\bсогласно\s+([А-Яа-яЁё]+а)\b", re.IGNORECASE),
        "corrected": "согласно + дательный падеж",
        "type": "style",
        "description": "После «согласно» требуется дательный падеж: согласно приказу, письму, договору.",
    },
    {
        "pattern": re.compile(r"\b(короче|типа|окей|ребят[а-я]*|привет)\b", re.IGNORECASE),
        "corrected": "замените на нейтральную официально-деловую формулировку",
        "type": "style",
        "description": "Разговорная лексика не подходит для официально-делового письма.",
    },
)


@dataclass
class TextBlock:
    text: str
    kind: str = "paragraph"
    table_index: int | None = None
    row_index: int | None = None
    cell_index: int | None = None
    style: dict[str, Any] = field(default_factory=dict)


def get_llm_client():
    global _llm_client
    if _llm_client is None:
        import openai

        openai.api_base = LLM_BASE_URL
        openai.api_key = "EMPTY"
        _llm_client = openai
    return _llm_client


def _json_from_llm(raw: str) -> dict[str, Any]:
    raw = raw.strip()
    raw = re.sub(r"^```[a-zA-Z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw)


def _alignment(paragraph: Paragraph) -> str:
    align_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    return align_map.get(paragraph.alignment, "left")


def _paragraph_style(paragraph: Paragraph) -> dict[str, Any]:
    style = {
        "alignment": _alignment(paragraph),
        "bold": False,
        "indent_left": 0,
        "first_line_indent": 0,
        "line_spacing": None,
    }

    fmt = paragraph.paragraph_format
    if fmt.left_indent:
        style["indent_left"] = round(fmt.left_indent.pt, 1)
    if fmt.first_line_indent:
        style["first_line_indent"] = round(fmt.first_line_indent.pt, 1)
    if fmt.line_spacing:
        try:
            style["line_spacing"] = float(fmt.line_spacing)
        except TypeError:
            style["line_spacing"] = None

    style["bold"] = any(run.bold for run in paragraph.runs if run.text.strip())
    return style


def _iter_body_items(doc: Document):
    table_index = 0
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc), table_index
            table_index += 1


def _clean_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def _read_blocks(doc: Document) -> tuple[list[TextBlock], list[list[list[str]]]]:
    blocks: list[TextBlock] = []
    table_rows: list[list[list[str]]] = []

    for item in _iter_body_items(doc):
        if isinstance(item, tuple):
            table, table_index = item
            current_rows: list[list[str]] = []
            for row_index, row in enumerate(table.rows):
                row_cells: list[str] = []
                seen: set[str] = set()
                for cell_index, cell in enumerate(row.cells):
                    text = "\n".join(_clean_lines(cell.text))
                    if text in seen:
                        continue
                    seen.add(text)
                    row_cells.append(text)
                    if text:
                        blocks.append(TextBlock(
                            text=text,
                            kind="table_cell",
                            table_index=table_index,
                            row_index=row_index,
                            cell_index=cell_index,
                            style={"alignment": "right" if cell_index > 0 else "left"},
                        ))
                if any(row_cells):
                    current_rows.append(row_cells)
            table_rows.append(current_rows)
        else:
            paragraph = item
            text = paragraph.text.strip()
            blocks.append(TextBlock(
                text=text,
                kind="paragraph",
                style=_paragraph_style(paragraph),
            ))

    return blocks, table_rows


def _find_greeting_index(blocks: list[TextBlock]) -> int | None:
    for index, block in enumerate(blocks):
        if block.text and GREETING_RE.search(block.text):
            return index
    return None


def _compact_header_rows(blocks: list[TextBlock], table_rows: list[list[list[str]]]) -> list[list[str]]:
    table_blocks = [block for block in blocks if block.kind == "table_cell" and block.text]
    if not table_blocks:
        return [[block.text] for block in blocks if block.text]

    rows: list[list[str]] = []
    current_key: tuple[int | None, int | None] | None = None
    current_cells: list[str] = []
    for block in sorted(table_blocks, key=lambda b: (b.table_index or 0, b.row_index or 0, b.cell_index or 0)):
        key = (block.table_index, block.row_index)
        if current_key is not None and key != current_key:
            rows.append(current_cells)
            current_cells = []
        current_key = key
        current_cells.append(block.text)
    if current_cells:
        rows.append(current_cells)
    return rows


def _score_addressee_line(line: str) -> int:
    lower = line.lower()
    score = 0
    if FIO_SURNAME_INITIALS_RE.search(line) or FIO_INITIALS_SURNAME_RE.search(line) or FULL_NAME_RE.search(line):
        score += 6
    if any(word in lower for word in ("кому", "адресат", "директор", "начальник", "руководител", "суд", "прокурор")):
        score += 3
    if re.search(r"\b(г\.|город|ул\.|улица|обл\.|район|р-н)\b", lower):
        score += 1
    if len(line) > 140:
        score -= 2
    return score


def _extract_addressee(header_blocks: list[TextBlock]) -> str | None:
    if not header_blocks:
        return None

    table_cells = [b for b in header_blocks if b.kind == "table_cell" and b.text]
    if table_cells:
        by_cell = sorted(table_cells, key=lambda b: (b.table_index or 0, b.row_index or 0, b.cell_index or 0))
        right_cells = [b for b in by_cell if (b.cell_index or 0) > 0]
        if not right_cells:
            return None
        candidates = right_cells
        best_index = max(range(len(candidates)), key=lambda i: _score_addressee_line(candidates[i].text))
        start = max(0, best_index - 1)
        end = min(len(candidates), best_index + 3)
        lines = []
        for block in candidates[start:end]:
            lines.extend(_clean_lines(block.text))
        return "\n".join(dict.fromkeys(lines)) or None

    non_empty = [b.text for b in header_blocks if b.text]
    scored = [(i, _score_addressee_line(text)) for i, text in enumerate(non_empty)]
    scored = [item for item in scored if item[1] > 0]
    if not scored:
        return "\n".join(non_empty[-4:]) if non_empty else None
    best = max(scored, key=lambda item: item[1])[0]
    return "\n".join(non_empty[max(0, best - 1): best + 3])


def extract_fio_list(addressee_text: str | None) -> list[dict[str, Any]]:
    if not addressee_text:
        return []

    people: list[dict[str, Any]] = []
    seen: set[tuple[str | None, str | None, str | None]] = set()

    for surname, i1, i2 in FIO_SURNAME_INITIALS_RE.findall(addressee_text):
        key = (surname.lower(), i1, i2 or None)
        if key not in seen:
            seen.add(key)
            people.append({
                "raw": f"{surname} {i1}.{(i2 + '.') if i2 else ''}".strip(),
                "surname": surname,
                "name": None,
                "patronymic": None,
                "i1": i1,
                "i2": i2 or None,
            })

    for i1, i2, surname in FIO_INITIALS_SURNAME_RE.findall(addressee_text):
        key = (surname.lower(), i1, i2 or None)
        if key not in seen:
            seen.add(key)
            people.append({
                "raw": f"{i1}.{(i2 + '.') if i2 else ''} {surname}".strip(),
                "surname": surname,
                "name": None,
                "patronymic": None,
                "i1": i1,
                "i2": i2 or None,
            })

    for surname, name, patronymic in FULL_NAME_RE.findall(addressee_text):
        if surname.lower() in {"уважаемый", "уважаемая", "уважаемые"}:
            continue
        key = (surname.lower(), name[:1], patronymic[:1])
        if key not in seen:
            seen.add(key)
            people.append({
                "raw": f"{surname} {name} {patronymic}",
                "surname": surname,
                "name": name,
                "patronymic": patronymic,
                "i1": name[:1],
                "i2": patronymic[:1],
            })

    return people


def parse_greeting(greeting_text: str | None) -> dict[str, Any]:
    if not greeting_text:
        return {"salutation": None, "is_collective": False, "persons": []}

    match = GREETING_RE.search(greeting_text.strip())
    salutation = match.group(1).split()[0].strip("!, .") if match else None
    lower = greeting_text.lower()
    if any(word in lower for word in COLLECTIVE_WORDS):
        return {"salutation": salutation, "is_collective": True, "persons": []}

    tail = re.sub(r"^\s*Уважаем(?:ый|ая|ые|ое)\s+", "", greeting_text, flags=re.IGNORECASE)
    tail = tail.strip("!,. \n\t")
    parts = [p.strip() for p in re.split(r"\s*,\s*|\s+и\s+", tail) if p.strip()]

    persons = []
    for part in parts:
        words = re.findall(r"[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?", part)
        if not words:
            continue
        if len(words) >= 3:
            surname, name, patronymic = words[-3], words[-2], words[-1]
        elif len(words) == 2:
            surname = None
            name, patronymic = words
        else:
            surname = None
            name, patronymic = words[0], None
        persons.append({
            "surname": surname,
            "name": name,
            "patronymic": patronymic,
            "i1": name[:1],
            "i2": patronymic[:1] if patronymic else None,
        })

    return {"salutation": salutation, "is_collective": not persons, "persons": persons}


def _gender_from_person(person: dict[str, Any]) -> str | None:
    patronymic = (person.get("patronymic") or "").lower()
    name = (person.get("name") or "").lower()
    if patronymic.endswith(MALE_PATRONYMIC_ENDINGS):
        return "male"
    if patronymic.endswith(FEMALE_PATRONYMIC_ENDINGS):
        return "female"
    if name in COMMON_MALE_NAMES:
        return "male"
    if name in COMMON_FEMALE_NAMES:
        return "female"
    return None


def _expected_salutation(parsed: dict[str, Any], fio_list: list[dict[str, Any]]) -> str | None:
    if parsed.get("is_collective") or len(fio_list) > 1 or len(parsed.get("persons", [])) > 1:
        return "Уважаемые"
    person = (parsed.get("persons") or [{}])[0]
    gender = _gender_from_person(person)
    if gender == "male":
        return "Уважаемый"
    if gender == "female":
        return "Уважаемая"
    return None


def check_gender(greeting_text: str | None, fio_list: list[dict[str, Any]], parsed: dict[str, Any]) -> dict[str, Any]:
    if not greeting_text:
        return {"status": "skip", "details": "Обращение отсутствует"}
    salutation = parsed.get("salutation")
    if not salutation:
        return {"status": "skip", "details": "Слово обращения не распознано"}

    expected = _expected_salutation(parsed, fio_list)
    if expected and salutation.lower() != expected.lower():
        return {
            "status": "error_gender",
            "details": f"В обращении указано «{salutation}», ожидается «{expected}».",
            "corrected": expected,
            "errors": [{
                "title": "Ошибка в роде или числе обращения",
                "description": f"Слово «{salutation}» не согласовано с адресатом.",
                "corrected": expected,
            }],
        }
    if expected:
        return {"status": "ok", "details": f"Обращение «{salutation}» согласовано с адресатом"}

    return check_gender_llm(greeting_text, fio_list, parsed)


def check_gender_llm(greeting_text: str, fio_list: list[dict[str, Any]], parsed: dict[str, Any]) -> dict[str, Any]:
    if os.getenv("LETTERCHECK_DISABLE_LLM") == "1":
        return {"status": "skip", "details": "LLM отключена настройкой окружения"}

    prompt = f"""Ты проверяешь русское деловое письмо.

Адресаты из шапки: {json.dumps(fio_list, ensure_ascii=False)}
Обращение: {greeting_text}
Разбор обращения: {json.dumps(parsed, ensure_ascii=False)}

Проверь только слово обращения: «Уважаемый», «Уважаемая» или «Уважаемые».
Верни строго JSON:
{{"correct": true}}
или
{{"correct": false, "reason": "кратко", "corrected": "Уважаемый/Уважаемая/Уважаемые"}}
"""
    try:
        client = get_llm_client()
        resp = client.ChatCompletion.create(
            model=LLM_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=200,
        )
        data = _json_from_llm(resp.choices[0].message.content)
        if data.get("correct"):
            return {"status": "ok", "details": "Род и число обращения выглядят корректно"}
        corrected = data.get("corrected") or "Уважаемый/Уважаемая"
        return {
            "status": "error_gender",
            "details": data.get("reason", "Обращение не согласовано с адресатом"),
            "corrected": corrected,
            "errors": [{
                "title": "Ошибка в роде обращения",
                "description": data.get("reason", "Обращение не согласовано с адресатом"),
                "corrected": corrected,
            }],
        }
    except Exception as exc:
        return {"status": "llm_error", "details": f"Не удалось проверить род через LLM: {exc}"}


def _compare_one(fio: dict[str, Any], person: dict[str, Any]) -> list[dict[str, Any]]:
    errors: list[dict[str, Any]] = []
    if (
        fio.get("surname")
        and person.get("surname")
        and fio["surname"].lower() != person["surname"].lower()
    ):
        errors.append({
            "title": "Ошибка в фамилии",
            "description": f"В шапке указана фамилия «{fio['surname']}», в обращении — «{person['surname']}».",
            "corrected": f"Используйте фамилию «{fio['surname']}» либо уберите фамилию из обращения.",
        })
    if fio.get("i1") and person.get("i1") and fio["i1"].upper() != person["i1"].upper():
        errors.append({
            "title": "Ошибка в имени",
            "description": f"В шапке имя начинается на «{fio['i1']}», в обращении — на «{person['i1']}».",
            "corrected": f"Имя должно начинаться на «{fio['i1']}».",
        })
    if fio.get("i2") and person.get("i2") and fio["i2"].upper() != person["i2"].upper():
        errors.append({
            "title": "Ошибка в отчестве",
            "description": f"В шапке отчество начинается на «{fio['i2']}», в обращении — на «{person['i2']}».",
            "corrected": f"Отчество должно начинаться на «{fio['i2']}».",
        })
    if fio.get("i2") and not person.get("i2"):
        errors.append({
            "title": "Не указано отчество",
            "description": f"В шапке есть второй инициал «{fio['i2']}», а в обращении отчество отсутствует.",
            "corrected": f"Добавьте отчество на «{fio['i2']}».",
        })
    return errors


def check_initials(fio_list: list[dict[str, Any]], greeting: dict[str, Any]) -> dict[str, Any]:
    persons = greeting.get("persons", [])
    is_collective = greeting.get("is_collective", False)

    if not fio_list:
        return {"status": "ok_no_fio", "details": "Адресат без распознанного физического лица"}
    if len(fio_list) > 2 and (persons or is_collective):
        return {
            "status": "error_too_many_addressees_greeting",
            "details": (
                "В шапке больше двух адресатов. По инструкции в таком случае "
                "вступительное обращение не указывается."
            ),
            "errors": [{
                "title": "Обращение при большом числе адресатов",
                "description": "Если адресатов больше двух, обращение к адресатам следует убрать.",
                "corrected": "Удалите вступительное обращение.",
            }],
        }
    if is_collective:
        if len(fio_list) == 1:
            return {
                "status": "error_personal_to_multiple",
                "details": "В шапке один адресат, но обращение коллективное.",
                "errors": [{
                    "title": "Личное обращение ожидается",
                    "description": "Для одного адресата лучше использовать личное обращение по имени и отчеству.",
                    "corrected": "Уважаемый/Уважаемая Имя Отчество!",
                }],
            }
        return {"status": "ok_collective", "details": "Коллективное обращение согласовано с несколькими адресатами"}
    if not persons:
        return {"status": "ok_no_greeting", "details": "Личное обращение не найдено"}
    if len(fio_list) != len(persons):
        return {
            "status": "error_mismatch",
            "details": f"В шапке адресатов: {len(fio_list)}, в обращении: {len(persons)}.",
            "errors": [{
                "title": "Разное количество адресатов",
                "description": f"В шапке адресатов: {len(fio_list)}, в обращении: {len(persons)}.",
                "corrected": "Согласуйте количество адресатов в шапке и обращении.",
            }],
        }

    all_errors: list[dict[str, Any]] = []
    for index, (fio, person) in enumerate(zip(fio_list, persons), 1):
        for error in _compare_one(fio, person):
            all_errors.append({
                **error,
                "description": f"Адресат {index} ({fio['raw']}): {error['description']}",
            })
    if all_errors:
        return {
            "status": "error_mismatch",
            "details": "; ".join(error["description"] for error in all_errors),
            "errors": all_errors,
        }
    return {"status": "ok", "details": "Фамилия и инициалы в шапке и обращении согласованы"}


def _line_offsets(text: str, phrase: str) -> tuple[int, int] | None:
    index = text.find(phrase)
    if index == -1:
        index = text.lower().find(phrase.lower())
    if index == -1:
        return None
    return index, index + len(phrase)


def check_normative_links(body_text: str) -> list[dict[str, Any]]:
    errors: list[dict[str, Any]] = []
    for match in NORMATIVE_RE.finditer(body_text):
        fragment = match.group(1).strip()
        lower = fragment.lower()
        if len(fragment) < 8:
            continue
        code_abbreviation = re.search(r"\b(?:гк|гпк|апк|кас|коап)\s+рф\b", lower)
        named_act = any(word in lower for word in ("закон", "постановлен", "приказ", "распоряжен"))
        if code_abbreviation and not named_act:
            continue
        has_date = bool(re.search(r"\bот\s+\d{1,2}(?:\.\d{1,2}\.\d{4}|\s+[а-яё]+\s+\d{4}\s+г(?:ода)?\.?)", lower))
        has_number = bool(re.search(r"№\s*[\wА-Яа-яЁё/-]+", fragment))
        needs_quotes = named_act
        has_quotes = bool(re.search(r"[«\"].+?[»\"]", fragment))
        missing = []
        if not has_date:
            missing.append("дата")
        if not has_number:
            missing.append("номер")
        if needs_quotes and not has_quotes:
            missing.append("название в кавычках")
        if missing:
            errors.append({
                "original": fragment,
                "corrected": "Укажите нормативную ссылку в формате: вид акта от ДД.ММ.ГГГГ № НОМЕР «Название».",
                "type": "normative",
                "description": "В нормативной ссылке не хватает: " + ", ".join(missing) + ".",
                "start": match.start(1),
                "end": match.end(1),
            })
    return errors


def _is_url_or_email_context(text: str, start: int, end: int) -> bool:
    window = text[max(0, start - 32): min(len(text), end + 32)].lower()
    return "http://" in window or "https://" in window or "@" in window


def check_deterministic_text_rules(text: str) -> list[dict[str, Any]]:
    """Fast local checks that do not depend on LLM availability."""
    errors: list[dict[str, Any]] = []
    used_spans: set[tuple[int, int, str]] = set()

    def add_error(start: int, end: int, original: str, corrected: str, error_type: str, description: str) -> None:
        key = (start, end, error_type)
        if key in used_spans:
            return
        used_spans.add(key)
        errors.append({
            "original": original,
            "corrected": corrected,
            "type": error_type,
            "description": description,
            "start": start,
            "end": end,
        })

    for rule in TEXT_RULES:
        for match in rule["pattern"].finditer(text):
            add_error(
                match.start(),
                match.end(),
                match.group(0),
                rule["corrected"],
                rule["type"],
                rule["description"],
            )

    for match in re.finditer(r"\s+([,;:])", text):
        if "\n" in match.group(0):
            continue
        add_error(
            match.start(),
            match.end(),
            match.group(0),
            match.group(1),
            "punctuation",
            "Перед запятой, точкой с запятой или двоеточием пробел не ставится.",
        )

    for match in re.finditer(r"(?<!\d)([,;:])(?=[А-ЯЁA-Za-zа-яё])", text):
        if _is_url_or_email_context(text, match.start(), match.end()):
            continue
        add_error(
            match.start(),
            match.end() + 1,
            text[match.start():match.end() + 1],
            match.group(1) + " " + text[match.end()],
            "punctuation",
            "После запятой, точки с запятой или двоеточия нужен пробел.",
        )

    for match in re.finditer(r"\.{2,}\s*\.", text):
        add_error(
            match.start(),
            match.end(),
            match.group(0),
            "…",
            "punctuation",
            "В тексте найдено избыточное многоточие или лишняя точка.",
        )

    return errors


def check_text_llm(text: str) -> dict[str, Any]:
    if not text.strip():
        return {"status": "skip", "details": "Нет текста для проверки", "errors": []}
    if os.getenv("LETTERCHECK_DISABLE_LLM") == "1":
        return {"status": "ok", "details": "LLM отключена настройкой окружения", "errors": []}

    prompt = f"""Ты корректор официально-делового русского текста.

Проверь текст на:
1. орфографию;
2. пунктуацию;
3. стиль официально-деловой речи.

Не переписывай весь текст. Верни только конкретные найденные проблемы.

Текст:
\"\"\"{text}\"\"\"

Ответ строго JSON без markdown:
{{
  "errors": [
    {{
      "original": "точный фрагмент из текста",
      "corrected": "исправленный фрагмент",
      "type": "spelling | punctuation | style",
      "description": "краткое объяснение"
    }}
  ]
}}
Если ошибок нет: {{"errors": []}}
"""
    try:
        client = get_llm_client()
        resp = client.ChatCompletion.create(
            model=LLM_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0,
            max_tokens=2200,
        )
        data = _json_from_llm(resp.choices[0].message.content)
        positioned = []
        for error in data.get("errors", []):
            original = (error.get("original") or "").strip()
            if not original:
                continue
            pos = _line_offsets(text, original)
            if not pos:
                continue
            start, end = pos
            positioned.append({
                "original": original,
                "corrected": error.get("corrected", ""),
                "type": error.get("type", "spelling"),
                "description": error.get("description", ""),
                "start": start,
                "end": end,
            })
        return {
            "status": "ok" if not positioned else "has_errors",
            "details": "Ошибок не найдено" if not positioned else f"Найдено замечаний: {len(positioned)}",
            "errors": positioned,
        }
    except Exception as exc:
        return {"status": "llm_error", "details": f"Не удалось проверить текст через LLM: {exc}", "errors": []}


def check_main_text(text: str) -> dict[str, Any]:
    normative_errors = check_normative_links(text)
    deterministic_errors = check_deterministic_text_rules(text)
    llm_result = check_text_llm(text)
    llm_errors = llm_result.get("errors", [])
    errors = normative_errors + deterministic_errors + llm_errors
    if errors:
        return {
            "status": "has_errors",
            "details": f"Найдено замечаний: {len(errors)}",
            "errors": errors,
        }
    if llm_result.get("status") == "llm_error":
        return {
            "status": "llm_error" if not normative_errors else "has_errors",
            "details": llm_result.get("details", "LLM недоступна"),
            "errors": normative_errors,
        }
    return {"status": "ok", "details": "Ошибок не найдено", "errors": []}


def classify_document(file_name: str, blocks: dict[str, Any]) -> dict[str, Any]:
    lower_name = file_name.lower()
    full_text = (blocks.get("full_text") or "").lower()
    header_text = "\n".join(" ".join(row) for row in blocks.get("header_rows", [])).lower()
    court_re = re.compile(r"\b(?:суд|суда|суду|арбитражн\w*)\b", re.IGNORECASE)

    if court_re.search(lower_name) or court_re.search(header_text):
        return {
            "type": "court_letter",
            "label": DOCUMENT_TYPE_LABELS["court_letter"],
            "confidence": "high",
            "reasons": ["В имени файла или адресате найден признак суда"],
        }
    if (
        "без обращения" in lower_name
        or "тема" in lower_name
        or "общий" in lower_name
        or "приложение на бумаге" in lower_name
    ):
        return {
            "type": "outgoing_without_greeting",
            "label": DOCUMENT_TYPE_LABELS["outgoing_without_greeting"],
            "confidence": "medium",
            "reasons": ["По имени шаблона выбран вариант исходящего письма без обращения"],
        }
    if blocks.get("greeting"):
        return {
            "type": "outgoing_letter",
            "label": DOCUMENT_TYPE_LABELS["outgoing_letter"],
            "confidence": "high",
            "reasons": ["В документе найдено обращение"],
        }
    if "уважаем" in full_text:
        return {
            "type": "outgoing_letter",
            "label": DOCUMENT_TYPE_LABELS["outgoing_letter"],
            "confidence": "medium",
            "reasons": ["В тексте найдено слово обращения"],
        }
    return {
        "type": "outgoing_letter",
        "label": DOCUMENT_TYPE_LABELS["outgoing_letter"],
        "confidence": "low",
        "reasons": ["Тип не распознан по сильным признакам, применен базовый шаблон исходящего письма"],
    }


def classify_document_type(file_name: str, blocks: dict[str, Any]) -> str:
    return classify_document(file_name, blocks)["type"]


def _preview_text(text: str | None, limit: int = 180) -> str | None:
    normalized = " ".join((text or "").split())
    if not normalized:
        return None
    if len(normalized) <= limit:
        return normalized
    return normalized[:limit - 1].rstrip() + "…"


def _find_attachment_evidence(full_text: str) -> str | None:
    for line in _clean_lines(full_text):
        if "приложение:" in line.lower():
            return line
    return None


def build_block_evidence(blocks: dict[str, Any]) -> dict[str, Any]:
    header_text = "\n".join(" | ".join(row) for row in blocks.get("header_rows", []))
    body_text = "\n".join(
        paragraph.get("text", "")
        for paragraph in blocks.get("body_paragraphs", [])
        if paragraph.get("text")
    )
    signature_text = "\n".join(
        paragraph.get("text", "")
        for paragraph in blocks.get("signature_paragraphs", [])
        if paragraph.get("text")
    )
    executor_text = "\n".join(
        paragraph.get("text", "")
        for paragraph in blocks.get("executor_paragraphs", [])
        if paragraph.get("text")
    )
    full_text = blocks.get("full_text") or ""

    return {
        "has_header": _preview_text(header_text),
        "has_addressee": _preview_text(blocks.get("addressee")),
        "has_greeting": _preview_text(blocks.get("greeting")),
        "has_body": _preview_text(body_text),
        "has_signature": _preview_text(signature_text),
        "has_executor": _preview_text(executor_text),
        "has_attachment": _preview_text(_find_attachment_evidence(full_text)),
    }


def check_required_blocks(file_name: str, blocks: dict[str, Any]) -> dict[str, Any]:
    classification = classify_document(file_name, blocks)
    doc_type = classification["type"]
    detected = blocks.get("detected", {})
    body_text = blocks.get("body_text") or ""
    full_text = blocks.get("full_text") or ""
    lower_name = file_name.lower()

    required = {
        "has_header": "шапка",
        "has_addressee": "адресат",
        "has_body": "основной текст",
        "has_signature": "подпись",
        "has_executor": "отметка об исполнителе",
    }
    if doc_type == "outgoing_letter":
        required["has_greeting"] = "обращение"
    if "приложение" in lower_name or "приложение:" in full_text.lower():
        required["has_attachment"] = "отметка о приложении"

    actual = {
        "has_header": bool(detected.get("has_header")),
        "has_addressee": bool(blocks.get("addressee")),
        "has_body": bool(body_text.strip()),
        "has_signature": bool(detected.get("has_signature")),
        "has_executor": bool(detected.get("has_executor")),
        "has_greeting": bool(detected.get("has_greeting")),
        "has_attachment": "приложение:" in full_text.lower(),
    }
    evidence = build_block_evidence(blocks)

    errors = []
    for key, label in required.items():
        if not actual.get(key):
            errors.append({
                "title": f"Отсутствует блок: {label}",
                "description": f"Для типа документа «{doc_type}» обязателен блок «{label}».",
                "block": key,
            })

    block_rows = []
    for key in BLOCK_ORDER:
        is_required = key in required
        is_present = bool(actual.get(key))
        if is_required and not is_present:
            status = "missing"
        elif is_required:
            status = "ok"
        elif is_present:
            status = "optional_present"
        else:
            status = "optional_absent"
        block_rows.append({
            "key": key,
            "label": BLOCK_LABELS[key],
            "required": is_required,
            "present": is_present,
            "status": status,
            "evidence": evidence.get(key),
        })

    return {
        "status": "ok" if not errors else "has_errors",
        "details": "Все обязательные блоки на месте" if not errors else f"Не хватает блоков: {len(errors)}",
        "document_type": doc_type,
        "document_type_label": classification["label"],
        "classification": classification,
        "required": required,
        "actual": actual,
        "blocks": block_rows,
        "errors": errors,
    }


def check_layout(blocks: dict[str, Any]) -> dict[str, Any]:
    errors = []
    paragraphs = (
        blocks.get("body_paragraphs", [])
        + blocks.get("signature_paragraphs", [])
        + blocks.get("executor_paragraphs", [])
    )

    for index, paragraph in enumerate(paragraphs, 1):
        if paragraph.get("empty") or not paragraph.get("text"):
            continue
        style = paragraph.get("style") or {}
        line_spacing = style.get("line_spacing")
        if line_spacing is not None and (line_spacing < 0.9 or line_spacing > 1.5):
            errors.append({
                "title": "Неверный межстрочный интервал",
                "description": (
                    f"Абзац {index}: межстрочный интервал {line_spacing}; "
                    "для исходящего письма ожидается обычный одинарный интервал."
                ),
                "type": "line_spacing",
                "paragraph": index,
                "value": line_spacing,
            })

    return {
        "status": "ok" if not errors else "has_errors",
        "details": "Интервалы выглядят корректно" if not errors else f"Найдено ошибок интервала: {len(errors)}",
        "errors": errors,
    }


def _result_kind(status: str | None) -> str:
    if not status:
        return "warn"
    if status.startswith("ok"):
        return "ok"
    if status.startswith("error") or status == "has_errors":
        return "error"
    if status == "skip":
        return "skip"
    return "warn"


def _check_group(name: str, task: str, checks: list[dict[str, Any]]) -> dict[str, Any]:
    kinds = [_result_kind(check.get("status")) for check in checks]
    if any(kind == "error" for kind in kinds):
        status = "has_errors"
    elif any(kind == "warn" for kind in kinds):
        status = "needs_attention"
    elif all(kind == "skip" for kind in kinds):
        status = "skip"
    else:
        status = "ok"

    errors = []
    for check in checks:
        errors.extend(check.get("errors", []) or [])

    return {
        "task": task,
        "name": name,
        "status": status,
        "details": "Проверка пройдена" if status == "ok" else "Есть замечания",
        "errors_count": len(errors),
        "checks": [
            {
                "status": check.get("status"),
                "details": check.get("details"),
            }
            for check in checks
        ],
    }


def build_quality_summary(
    gender_result: dict[str, Any],
    initials_result: dict[str, Any],
    text_result: dict[str, Any],
    required_blocks_result: dict[str, Any],
) -> dict[str, Any]:
    groups = [
        _check_group(
            "Адресат и обращение",
            "Задача 1",
            [gender_result, initials_result],
        ),
        _check_group(
            "Основной текст и нормативные ссылки",
            "Задача 2",
            [text_result],
        ),
        _check_group(
            "Обязательные блоки документа",
            "Задача 3",
            [required_blocks_result],
        ),
    ]

    error_groups = [group for group in groups if group["status"] == "has_errors"]
    attention_groups = [group for group in groups if group["status"] == "needs_attention"]
    total_errors = sum(group["errors_count"] for group in groups)
    score = max(0, 100 - len(error_groups) * 24 - total_errors * 4 - len(attention_groups) * 8)

    if error_groups:
        status = "needs_revision"
        verdict = "Документ требует исправлений перед отправкой."
    elif attention_groups:
        status = "needs_attention"
        verdict = "Критичных ошибок нет, но есть проверки, требующие внимания."
    else:
        status = "ready"
        verdict = "Документ готов к отправке по проверяемым критериям."

    recommendations = []
    for group in groups:
        if group["status"] == "has_errors":
            recommendations.append(f"Исправить раздел «{group['name']}».")
    if not recommendations:
        recommendations.append("Сохранить текущую структуру письма; проверяемые требования выполнены.")

    return {
        "status": status,
        "score": score,
        "verdict": verdict,
        "total_errors": total_errors,
        "groups": groups,
        "recommendations": recommendations,
    }


def _split_signature(body_blocks: list[TextBlock]) -> tuple[list[TextBlock], list[TextBlock], list[TextBlock]]:
    non_empty_indexes = [i for i, block in enumerate(body_blocks) if block.text]
    if not non_empty_indexes:
        return [], [], []

    executor_start = None
    for i in non_empty_indexes:
        text = body_blocks[i].text.strip()
        lower = text.lower()
        is_executor_like = bool(EXECUTOR_RE.search(text)) and (
            len(text) <= 160
            or lower.startswith(("исп", "тел", "8 ("))
            or re.fullmatch(r"[\d\s()+\-.]+(?:\(доб\.?\s*\d+\))?", text)
        )
        if is_executor_like:
            executor_start = i
            break

    usable = body_blocks[:executor_start] if executor_start is not None else body_blocks
    executor = body_blocks[executor_start:] if executor_start is not None else []

    signature_start = None
    for i in range(len(usable) - 1, -1, -1):
        text = usable[i].text
        if not text:
            continue
        if SIGNATURE_RE.search(text) and len(text) <= 220:
            signature_start = max(0, i - 1 if i > 0 and not usable[i - 1].text else i)
            break

    if signature_start is None:
        return usable, [], executor
    return usable[:signature_start], [b for b in usable[signature_start:] if b.text], executor


def extract_document_blocks(docx_path: str) -> dict[str, Any]:
    doc = Document(docx_path)
    blocks, table_rows = _read_blocks(doc)
    greeting_index = _find_greeting_index(blocks)

    if greeting_index is None:
        header_blocks = []
        greeting_block = None
        body_source = blocks
    else:
        header_blocks = [b for b in blocks[:greeting_index] if b.text]
        greeting_block = blocks[greeting_index]
        body_source = blocks[greeting_index + 1:]

    if greeting_index is None:
        for index, block in enumerate(blocks):
            if BODY_START_RE.search(block.text):
                header_blocks = [b for b in blocks[:index] if b.text]
                body_source = blocks[index:]
                break

    body_blocks, signature_blocks, executor_blocks = _split_signature(body_source)
    body_paragraphs = [
        {"text": block.text, "style": block.style, "empty": not bool(block.text)}
        for block in body_blocks
    ]
    signature_paragraphs = [
        {"text": block.text, "style": block.style, "empty": False}
        for block in signature_blocks
    ]
    executor_paragraphs = [
        {"text": block.text, "style": block.style, "empty": False}
        for block in executor_blocks if block.text
    ]

    body_text = "\n".join(block.text for block in body_blocks if block.text)
    full_text = "\n".join(block.text for block in blocks if block.text)

    return {
        "header_rows": _compact_header_rows(header_blocks, table_rows),
        "addressee": _extract_addressee(header_blocks),
        "greeting": greeting_block.text if greeting_block else None,
        "body_paragraphs": body_paragraphs,
        "signature_paragraphs": signature_paragraphs,
        "executor_paragraphs": executor_paragraphs,
        "body_text": body_text,
        "full_text": full_text,
        "detected": {
            "has_header": bool(header_blocks),
            "has_greeting": greeting_block is not None,
            "has_signature": bool(signature_blocks),
            "has_executor": bool(executor_blocks),
        },
    }


def extract_header_and_greeting(docx_path: str) -> dict[str, Any]:
    blocks = extract_document_blocks(docx_path)
    file_name = Path(docx_path).name
    addressee = blocks.get("addressee")
    fio = extract_fio_list(addressee)
    greeting_text = blocks.get("greeting")
    parsed = parse_greeting(greeting_text)

    gender_result = check_gender(greeting_text, fio, parsed)
    initials_result = check_initials(fio, parsed)
    text_result = check_main_text(blocks.get("body_text") or blocks.get("full_text") or "")
    required_blocks_result = check_required_blocks(file_name, blocks)
    layout_result = check_layout(blocks)
    quality_summary = build_quality_summary(
        gender_result,
        initials_result,
        text_result,
        required_blocks_result,
    )

    greeting_errors = (
        gender_result.get("status", "").startswith("error")
        or initials_result.get("status", "").startswith("error")
    )

    return {
        "file": file_name,
        "mode": "docx",
        "blocks": blocks,
        "addressee": addressee,
        "fio": fio,
        "greeting": greeting_text,
        "parsed_greeting": parsed,
        "highlight_greeting": greeting_errors,
        "error_positions": [{
            "start": 0,
            "end": len(greeting_text or ""),
            "text": greeting_text,
        }] if greeting_text and greeting_errors else [],
        "check_gender": gender_result,
        "check_initials": initials_result,
        "check_spelling": text_result,
        "check_required_blocks": required_blocks_result,
        "check_layout": layout_result,
        "quality_summary": quality_summary,
    }


def check_plain_text(text: str) -> dict[str, Any]:
    normalized = text.strip()
    result = check_main_text(normalized)
    gender_result = {"status": "skip", "details": "Для обычного текста не применяется"}
    initials_result = {"status": "skip", "details": "Для обычного текста не применяется"}
    required_result = {"status": "skip", "details": "Для обычного текста не применяется", "errors": []}
    paragraphs = [
        {"text": line, "style": {"alignment": "justify"}, "empty": not bool(line.strip())}
        for line in normalized.splitlines()
    ] or [{"text": normalized, "style": {"alignment": "justify"}, "empty": False}]
    return {
        "file": "Вставленный текст",
        "mode": "text",
        "blocks": {
            "header_rows": [],
            "addressee": None,
            "greeting": None,
            "body_paragraphs": paragraphs,
            "signature_paragraphs": [],
            "executor_paragraphs": [],
            "body_text": normalized,
            "full_text": normalized,
            "detected": {
                "has_header": False,
                "has_greeting": False,
                "has_signature": False,
                "has_executor": False,
            },
        },
        "addressee": None,
        "fio": [],
        "greeting": None,
        "parsed_greeting": {"salutation": None, "is_collective": False, "persons": []},
        "highlight_greeting": False,
        "error_positions": [],
        "check_gender": gender_result,
        "check_initials": initials_result,
        "check_spelling": result,
        "check_required_blocks": required_result,
        "check_layout": {"status": "skip", "details": "Для обычного текста не применяется", "errors": []},
        "quality_summary": build_quality_summary(gender_result, initials_result, result, required_result),
    }
