import unittest
from copy import deepcopy
from pathlib import Path

from compare_contract_details import BANK_FIELDS, review_details
from inspect_contract_docx import SAMPLE_CONTRACTS, extract_blocks


def table_row(blocks, label):
    return next(block for block in blocks if block["kind"] == "table_row" and block["cells"][0] == label)


class ContractDetailsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template, cls.original_carrier = [
            list(extract_blocks(Path(__file__).parent / "examples" / name)) for name in SAMPLE_CONTRACTS
        ]

    def setUp(self):
        self.carrier = deepcopy(self.original_carrier)

    def review(self):
        return review_details(self.template, self.carrier)

    def only_issue(self, field, status):
        issues = [entry for entry in self.review() if entry["status"] != "совпадает"]
        self.assertEqual(
            [(entry["role"], entry["field"], entry["status"]) for entry in issues],
            [("Перевозчик", field, status)],
        )
        return issues[0]

    def test_original_contracts_match_with_both_name_sources(self):
        report = self.review()
        self.assertEqual(len(report), 10)
        self.assertTrue(all(entry["status"] == "совпадает" for entry in report))
        for entry in report:
            if entry["field"] == "Наименование":
                for side in ("template", "carrier"):
                    self.assertEqual(len(entry[side]), 2)
                    preamble, table = (record["source"] for record in entry[side])
                    self.assertEqual(preamble["headings"][-1], "Преамбула")
                    self.assertEqual((table["table"], table["row"]), (1, 2))
        tax_id = next(entry for entry in report if entry["role"] == "Перевозчик" and entry["field"] == "ИНН")
        self.assertEqual(tax_id["carrier"][0]["value"], "0000000001")

    def test_reordered_roles_and_expanded_address_abbreviations(self):
        address = table_row(self.carrier, "Юридический адрес")
        address["cells"][2] = "город Пермь, улица Макетная, дом 27, офис 12"
        for block in self.carrier:
            if block.get("table_number") == address["table_number"]:
                block["cells"][1], block["cells"][2] = block["cells"][2], block["cells"][1]
        self.assertTrue(all(entry["status"] == "совпадает" for entry in self.review()))

    def test_changed_legal_address_keeps_both_cell_references(self):
        table_row(self.carrier, "Юридический адрес")["cells"][2] = "г. Пермь, ул. Макетная, д. 27, офис 21"
        issue = self.only_issue("Юридический адрес", "отличается")
        for side, suffix in (("template", "офис 12"), ("carrier", "офис 21")):
            self.assertTrue(issue[side][0]["value"].endswith(suffix))
            source = issue[side][0]["source"]
            self.assertEqual((source["table"], source["row"], source["column"]), (1, 5, 3))

    def test_notice_and_unloading_addresses_are_separate(self):
        table_row(self.carrier, "Адрес для уведомлений")["cells"][2] = "г. Пермь, ул. Макетная, д. 27, офис 8"
        table_row(self.carrier, "Грузополучатель")["cells"][1] = "Адрес выгрузки: г. Ижевск, ул. Складская, д. 15"
        self.only_issue("Адрес для уведомлений", "отличается")

    def test_conflicting_address_in_appendix_keeps_both_sources(self):
        anchor = next(block for block in reversed(self.carrier) if block["kind"] == "table_row")
        for number, cells in enumerate((
            ["Реквизиты", "Перевозчик"],
            ["Юридический адрес", "г. Пермь, ул. Макетная, д. 27, офис 21"],
        ), start=1):
            self.carrier.append({
                **anchor, "table_number": 4, "body_position": anchor["body_position"] + 1,
                "row_number": number, "cells": cells, "text": "\t".join(cells),
            })
        issue = self.only_issue("Юридический адрес", "противоречивые значения")
        self.assertEqual(
            [(record["source"]["table"], record["source"]["row"], record["source"]["column"])
             for record in issue["carrier"]],
            [(1, 5, 3), (4, 2, 2)],
        )

    def test_missing_field_in_both_documents_requires_manual_review(self):
        template = deepcopy(self.template)
        for blocks in (template, self.carrier):
            table_row(blocks, "ИНН")["cells"][2] = ""
        report = review_details(template, self.carrier)
        issues = [entry for entry in report if entry["status"] != "совпадает"]
        self.assertEqual(len(issues), 1)
        self.assertEqual((issues[0]["role"], issues[0]["field"], issues[0]["status"]),
                         ("Перевозчик", "ИНН", "нужна ручная проверка"))
        self.assertEqual(set(issues[0]["reasons"]), {"template", "carrier"})

    def test_bank_rows_compare_aliases_and_one_sided_values(self):
        anchor = table_row(self.carrier, "Адрес для уведомлений")
        template = deepcopy(self.template)
        for blocks, values in (
            (template, [("БИК", "000 000 001"), ("Банк", "АО «Учебный банк»"), ("к/с", "001")]),
            (self.carrier, [
                ("БИК:", "000000001"), ("Банк получателя", "АО «Другой банк»"),
                ("р/с", "00000000000000000421"), ("Корреспондентский счет", "—"),
            ]),
        ):
            for number, (label, value) in enumerate(values, 20):
                blocks.append({**anchor, "row_number": number, "cells": [label, "—", value]})
        bank = {
            item["field"]: item for item in review_details(template, self.carrier)
            if item["role"] == "Перевозчик" and item["field"] in BANK_FIELDS
        }
        self.assertEqual({field: item["status"] for field, item in bank.items()}, {
            "БИК": "совпадает", "Банк": "отличается",
            "Расчётный счёт": "появилось в договоре", "Корреспондентский счёт": "не указано в договоре",
        })

    def test_separate_party_tables_use_heading_or_role_row(self):
        original = [block for block in self.carrier if block.get("table_number") == 1]
        self.carrier = [block for block in self.carrier if block.get("table_number") != 1]
        for column, role in enumerate(("Заказчик", "Перевозчик"), 1):
            for block in original:
                label, *values = block["cells"]
                if label == "Сведения":
                    if column == 2:
                        self.carrier.append({**block, "table_number": 12, "cells": [role, ""]})
                    continue
                self.carrier.append({
                    **block, "table_number": 10 + column, "cells": [label, values[column - 1]],
                    "headings": [role] if column == 1 else ["Реквизиты сторон"],
                })
        self.assertTrue(all(item["status"] == "совпадает" for item in self.review()))


if __name__ == "__main__":
    unittest.main()
