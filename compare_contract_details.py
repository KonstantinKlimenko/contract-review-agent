"""Compare party details extracted from two DOCX contracts."""

import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

from inspect_contract_docx import SAMPLE_CONTRACTS, extract_blocks


ROLES = ("Заказчик", "Перевозчик")
FIELDS = ("Наименование", "ИНН", "КПП", "Юридический адрес", "Адрес для уведомлений")
BANK_FIELDS = ("Расчётный счёт", "БИК", "Банк", "Корреспондентский счёт")
ROLE_LABELS = {role.casefold(): role for role in ROLES}
FIELD_LABELS = {field.casefold().replace("ё", "е"): field for field in FIELDS + BANK_FIELDS}
FIELD_LABELS.update({
    "р/с": "Расчётный счёт", "расчетный счет (р/с)": "Расчётный счёт",
    "к/с": "Корреспондентский счёт", "корр. счет": "Корреспондентский счёт",
    "наименование банка": "Банк", "банк получателя": "Банк",
})
ADDRESS_FIELDS = {"Юридический адрес", "Адрес для уведомлений"}
ADDRESS_WORDS = {"город": "г", "улица": "ул", "дом": "д", "корпус": "корп", "строение": "стр"}
ADDRESS_ABBREVIATION = re.compile(r"\b(г|ул|д|корп|стр)\.\s*")
PARTY_DECLARATION = re.compile(
    r'((?:ООО|АО|ПАО|ИП)\s+(?:«[^»]+»|"[^"]+"))\s*,\s*'
    r'именуем\w*\s+(?:далее\s+)?(Заказчиком|Перевозчиком)\b',
    re.IGNORECASE,
)
DECLARED_ROLES = {"заказчиком": "Заказчик", "перевозчиком": "Перевозчик"}


def fold(value):
    return " ".join(value.casefold().split())


def comparison_value(field, value):
    normalized = fold(value)
    if field in ADDRESS_FIELDS:
        normalized = ADDRESS_ABBREVIATION.sub(r"\1 ", normalized.replace("ё", "е"))
        words = normalized.replace(",", " ").split()
        return " ".join(ADDRESS_WORDS.get(word, word) for word in words)
    if field in ("ИНН", "КПП", BANK_FIELDS[0], BANK_FIELDS[1], BANK_FIELDS[3]):
        return re.sub(r"\s+", "", normalized)
    return normalized.translate(str.maketrans("", "", '«»"'))


def add_value(details, role, field, value, block, column=None):
    if value.strip() in ("", "-", "–", "—"):
        return
    source = {
        "document": block["source"],
        "headings": block["headings"],
        "body_position": block["body_position"],
        **{key: block[key] for key in ("page", "page_end") if key in block},
    }
    if column is not None:
        source.update(table=block["table_number"], row=block["row_number"], column=column)
    details[role, field].append({"value": value, "source": source})


def extract_details(blocks):
    """Use explicit role and field labels, retaining every source for conflict checks."""
    details = defaultdict(list)
    role_columns = {}
    for block in blocks:
        if block["kind"] == "table_row":
            cells = block["cells"]
            table = block["table_number"]
            label = fold(cells[0]).replace("ё", "е").rstrip(":")
            if label in ("сведения", "реквизиты"):
                role_columns[table] = [
                    (index, ROLE_LABELS[fold(cell)])
                    for index, cell in enumerate(cells[1:], 1)
                    if fold(cell) in ROLE_LABELS
                ]
            elif label in ROLE_LABELS and all(not cell.strip() for cell in cells[1:]):
                role_columns[table] = [(1, ROLE_LABELS[label])]
            elif label in FIELD_LABELS:
                if table not in role_columns and len(cells) == 2 and block["headings"]:
                    role = ROLE_LABELS.get(fold(block["headings"][-1]).rstrip(":"))
                    if role:
                        role_columns[table] = [(1, role)]
                for index, role in role_columns.get(table, ()):
                    if index < len(cells):
                        add_value(details, role, FIELD_LABELS[label], cells[index], block, index + 1)
        elif block["kind"] in ("paragraph", "clause"):
            for value, role in PARTY_DECLARATION.findall(block["text"]):
                add_value(details, DECLARED_ROLES[fold(role)], "Наименование", value, block)
    return details


def compare_field(role, field, template, carrier):
    template_values = {comparison_value(field, record["value"]) for record in template}
    carrier_values = {comparison_value(field, record["value"]) for record in carrier}
    reasons = {}
    for side, values in (("template", template_values), ("carrier", carrier_values)):
        if len(values) > 1:
            reasons[side] = "Найдены разные значения в одном документе"
        elif not values:
            reasons[side] = "Поле не найдено"

    if len(template_values) > 1 or len(carrier_values) > 1:
        status = "противоречивые значения"
    elif field in BANK_FIELDS and reasons:
        status = "появилось в договоре" if carrier_values else "не указано в договоре"
    elif reasons:
        status = "нужна ручная проверка"
    else:
        status = "совпадает" if template_values == carrier_values else "отличается"
    return {
        "role": role, "field": field, "status": status, "reasons": reasons,
        "template": template, "carrier": carrier,
    }


def review_details(template_blocks, carrier_blocks):
    template = extract_details(template_blocks)
    carrier = extract_details(carrier_blocks)
    return [
        compare_field(role, field, template.get((role, field), []), carrier.get((role, field), []))
        for role in ROLES
        for field in FIELDS + BANK_FIELDS
        if field in FIELDS or (role, field) in template or (role, field) in carrier
    ]


def main():
    parser = argparse.ArgumentParser(description="Сравнивает реквизиты заказчика и перевозчика в двух DOCX.")
    parser.add_argument("template", nargs="?", type=Path, default=Path(__file__).parent / "examples" / SAMPLE_CONTRACTS[0])
    parser.add_argument("carrier", nargs="?", type=Path, default=Path(__file__).parent / "examples" / SAMPLE_CONTRACTS[1])
    args = parser.parse_args()
    for path in (args.template, args.carrier):
        if path.suffix.lower() != ".docx" or not path.is_file():
            parser.error(f"Нужен существующий файл DOCX: {path}")

    report = review_details(extract_blocks(args.template), extract_blocks(args.carrier))
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
