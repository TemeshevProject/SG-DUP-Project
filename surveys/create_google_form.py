#!/usr/bin/env python3
"""Create Google Form from corrected DUP survey Excel.

Two modes:
  apps-script (default) — generates a .gs file for script.google.com (recommended)
  api                 — creates form via Google Forms API (needs OAuth credentials.json)

Usage:
  python3 create_google_form.py
  python3 create_google_form.py --mode api
  python3 create_google_form.py --xlsx "Опросник_директоров_Филиалов.xlsx"
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from parse_survey_xlsx import DEFAULT_XLSX, Survey, parse_survey_xlsx

SURVEYS_DIR = Path(__file__).resolve().parent
CREDENTIALS = SURVEYS_DIR / "credentials.json"
TOKEN = SURVEYS_DIR / "token.json"
OUTPUT_GS = SURVEYS_DIR / "create_DUP_google_form.gs"

SCOPES = [
    "https://www.googleapis.com/auth/forms.body",
    "https://www.googleapis.com/auth/drive.file",
]

SCALE_HINTS = {
    "понятен": ("1 — совсем непонятен", "5 — полностью понятен"),
    "полным приходит": ("1 — почти всегда неполный", "5 — почти всегда полный"),
    "сложно согласовать": ("1 — очень сложно", "5 — проходит быстро"),
    "своевременно": ("1 — очень медленно", "5 — оперативно"),
    "налажен процесс": ("1 — не налажен", "5 — отлично налажен"),
    "приоритизации": ("1 — совсем непонятен", "5 — полностью понятен"),
    "удобен текущий формат": ("1 — совсем неудобен", "5 — очень удобен"),
    "разграничены зоны": ("1 — совсем неясно", "5 — чётко разграничено"),
    "к кому в ДУП обращаться": ("1 — не понимаю", "5 — полностью понимаю"),
    "качество коммуникации": ("1 — очень плохо", "5 — отлично"),
    "ДРБ передаёт": ("1 — очень плохо", "5 — отлично"),
    "Юр блок": ("1 — очень медленно", "5 — оперативно"),
    "платформа портфеля": ("1 — не работает", "5 — работает полноценно"),
}


def scale_labels(text: str) -> tuple[str, str]:
    lower = text.lower()
    for key, labels in SCALE_HINTS.items():
        if key in lower:
            return labels
    return ("1 — плохо", "5 — хорошо")


def form_item_title(qid: str, text: str) -> str:
    return f"{qid}. {text}"


def js(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


def generate_apps_script(survey: Survey) -> str:
    lines: list[str] = [
        "/**",
        " * ДУП — создание Google Form из откорректированного опросника.",
        " *",
        " * Как запустить:",
        " * 1. Откройте https://script.google.com → Новый проект",
        " * 2. Вставьте этот код, сохраните",
        " * 3. Выберите функцию createDUPSurvey → Выполнить",
        " * 4. Разрешите доступ к Google Forms (при первом запуске)",
        " * 5. Ссылка на форму появится в журнале выполнения (View → Logs)",
        " */",
        "",
        "function createDUPSurvey() {",
        f"  var form = FormApp.create({js(survey.title)});",
        f"  form.setDescription({js(survey.description)});",
        "  form.setCollectEmail(true);",
        "  form.setLimitOneResponsePerUser(false);",
        "  form.setShuffleQuestions(false);",
        "",
    ]

    for section in survey.sections:
        lines.append(f"  form.addPageBreakItem().setTitle({js(section.title)});")
        lines.append("")

        for q in section.questions:
            title = form_item_title(q.qid, q.text)
            req = "true" if q.required else "false"

            if q.qtype == "text":
                lines.append(f"  form.addTextItem().setTitle({js(title)}).setRequired({req});")
            elif q.qtype == "paragraph":
                lines.append(
                    f"  form.addParagraphTextItem().setTitle({js(title)}).setRequired({req});"
                )
            elif q.qtype == "radio":
                choices = ", ".join(js(o) for o in q.options)
                lines.append(f"  form.addMultipleChoiceItem()")
                lines.append(f"    .setTitle({js(title)})")
                lines.append(f"    .setChoiceValues([{choices}])")
                lines.append(f"    .setRequired({req});")
            elif q.qtype == "checkbox":
                choices = ", ".join(js(o) for o in q.options)
                lines.append(f"  form.addCheckboxItem()")
                lines.append(f"    .setTitle({js(title)})")
                lines.append(f"    .setChoiceValues([{choices}])")
                lines.append(f"    .setRequired({req});")
            elif q.qtype == "scale":
                low, high = scale_labels(q.text)
                lines.append(f"  form.addScaleItem()")
                lines.append(f"    .setTitle({js(title)})")
                lines.append(f"    .setBounds(1, 5)")
                lines.append(f"    .setLabels({js(low)}, {js(high)})")
                lines.append(f"    .setRequired({req});")
            elif q.qtype == "grid":
                rows = ", ".join(js(o) for o in q.options)
                cols = ", ".join(js(str(i)) for i in range(1, 6))
                lines.append(f"  form.addGridItem()")
                lines.append(f"    .setTitle({js(title)})")
                lines.append(f"    .setRows([{rows}])")
                lines.append(f"    .setColumns([{cols}])")
                lines.append(f"    .setRequired({req});")
            lines.append("")

    lines.extend(
        [
            "  Logger.log('Форма создана: ' + form.getPublishedUrl());",
            "  Logger.log('Редактирование: ' + form.getEditUrl());",
            "}",
        ]
    )
    return "\n".join(lines)


def _choice_options(options: list[str]) -> list[dict]:
    return [{"value": o} for o in options]


def _build_api_requests(survey: Survey) -> list[dict]:
    requests: list[dict] = []
    index = 0

    for section in survey.sections:
        requests.append(
            {
                "createItem": {
                    "item": {
                        "title": section.title,
                        "pageBreakItem": {},
                    },
                    "location": {"index": index},
                }
            }
        )
        index += 1

        for q in section.questions:
            title = form_item_title(q.qid, q.text)
            item: dict = {"title": title}
            question: dict = {}

            if q.qtype == "text":
                question = {"textQuestion": {"paragraph": False}}
            elif q.qtype == "paragraph":
                question = {"textQuestion": {"paragraph": True}}
            elif q.qtype == "radio":
                question = {
                    "choiceQuestion": {
                        "type": "RADIO",
                        "options": _choice_options(q.options),
                        "shuffle": False,
                    }
                }
            elif q.qtype == "checkbox":
                question = {
                    "choiceQuestion": {
                        "type": "CHECKBOX",
                        "options": _choice_options(q.options),
                        "shuffle": False,
                    }
                }
            elif q.qtype == "scale":
                low, high = scale_labels(q.text)
                question = {
                    "scaleQuestion": {
                        "low": 1,
                        "high": 5,
                        "lowLabel": low,
                        "highLabel": high,
                    }
                }
            elif q.qtype == "grid":
                question = {
                    "rowQuestion": {
                        "grid": {
                            "columns": {
                                "type": "RADIO",
                                "options": _choice_options([str(i) for i in range(1, 6)]),
                            },
                            "rows": [{"value": r} for r in q.options],
                        }
                    }
                }

            item["questionItem"] = {"question": question, "required": q.required}
            requests.append(
                {"createItem": {"item": item, "location": {"index": index}}}
            )
            index += 1

    return requests


def create_via_api(survey: Survey) -> str:
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError:
        print("Установите: pip install google-api-python-client google-auth-oauthlib")
        sys.exit(1)

    if not CREDENTIALS.exists():
        print(f"Нужен OAuth credentials.json в {CREDENTIALS}")
        print("Создайте OAuth Client ID (Desktop) в Google Cloud Console.")
        sys.exit(1)

    creds = None
    if TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN.write_text(creds.to_json())

    service = build("forms", "v1", credentials=creds, cache_discovery=False)
    form = (
        service.forms()
        .create(body={"info": {"title": survey.title, "description": survey.description}})
        .execute()
    )
    form_id = form["formId"]
    requests = _build_api_requests(survey)
    service.forms().batchUpdate(formId=form_id, body={"requests": requests}).execute()
    return form.get("responderUri", f"https://docs.google.com/forms/d/{form_id}/viewform")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create Google Form from DUP survey Excel")
    parser.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX, help="Path to survey xlsx")
    parser.add_argument(
        "--mode",
        choices=("apps-script", "api"),
        default="apps-script",
        help="apps-script: generate .gs file; api: create form via OAuth",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_GS,
        help="Output path for Apps Script file",
    )
    args = parser.parse_args()

    if not args.xlsx.exists():
        print(f"Файл не найден: {args.xlsx}")
        sys.exit(1)

    survey = parse_survey_xlsx(args.xlsx)
    total = sum(len(s.questions) for s in survey.sections)
    print(f"Опросник: {len(survey.sections)} блоков, {total} вопросов")

    if args.mode == "apps-script":
        code = generate_apps_script(survey)
        args.output.write_text(code, encoding="utf-8")
        print(f"Сохранён Apps Script: {args.output}")
        print()
        print("Дальше:")
        print("  1. https://script.google.com → Новый проект")
        print(f"  2. Вставьте код из {args.output.name}")
        print("  3. Функция createDUPSurvey → Выполнить")
        print("  4. Ссылка на форму — в журнале (Журнал выполнения)")
    else:
        url = create_via_api(survey)
        print(f"Форма создана: {url}")


if __name__ == "__main__":
    main()
