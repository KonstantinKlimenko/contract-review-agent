import unittest
from io import BytesIO
from pathlib import Path

from docx import Document

from compare_contract_clauses import compare_clauses, decision_schema, read_clauses, template_fit, validate_decision
from inspect_contract_docx import pdf_blocks


common_a = {
    "1.1":"Перевозчик доставляет упакованные кабельные изделия между складами Заказчика по заявкам.",
    "1.2":"В каждой заявке стороны указывают номер партии, маршрут и окно подачи.",
    "2.1":"При задержке подачи более чем на 30 минут диспетчер сообщает Заказчику ожидаемое время прибытия.",
    "3.1":"Заказчик обеспечивает доступ автомобиля к зоне выгрузки в согласованное время.",
    "4.1":"В акте о повреждении стороны указывают номер партии и количество повреждённых мест.",
    "5.1":"Оплата оказанных услуг сохраняется при прекращении договора.",
}
left_a = {**common_a,
    "2.2":"На въезде на склад Заказчик организует сопровождение водителя до зоны выгрузки.",
    "3.2":"До закрытия кузова Перевозчик фотографирует установленную пломбу и записывает её номер в путевой лист.",
    "4.2":"Возражения по акту о повреждении Перевозчик направляет в течение трёх рабочих дней после его получения.",
}
right_a = {**common_a,
    "2.3":"После въезда на склад водитель следует в зону выгрузки в сопровождении, которое организует Заказчик.",
    "4.2":"Возражения по акту о повреждении Перевозчик направляет в течение трёх календарных дней после его получения.",
    "5.2":"Перевозчик вправе передавать контакты грузополучателей привлечённым исполнителям без отдельного согласия Заказчика.",
}
common_b = {
    "1.1":"Перевозчик выполняет сборные перевозки по согласованным маршрутам.",
    "1.2":"Заказчик указывает предельную массу и габариты отправления в заявке.",
    "2.1":"Диспетчер Перевозчика согласовывает с получателем время выдачи отправления.",
    "3.1":"Перевозчик передаёт груз в пункте назначения представителю, указанному в заявке.",
    "4.1":"Сумма за выполненную перевозку определяется по согласованной таблице маршрутов.",
    "5.1":"При разногласиях стороны обмениваются письменными претензиями.",
}
left_b = {**common_b,
    "2.2":"При приёмке Перевозчик проверяет целостность упаковки, число грузовых мест и состояние пломб.",
    "3.2":"До полудня каждого понедельника Перевозчик направляет Заказчику отчёт о пробеге машин за предыдущую неделю.",
    "4.2":"За нарушение окна подачи Перевозчик платит Заказчику 1 500 рублей за каждое нарушение.",
}
right_b = {**common_b,
    "2.3":"Принимая груз, Перевозчик сверяет количество мест и осматривает упаковку и пломбы.",
    "4.3":"За нарушение окна подачи Перевозчик платит Заказчику 1 500 рублей за каждый день, в который допущено одно или несколько нарушений.",
    "5.2":"При остановке на ночь Перевозчик вправе перегружать отправления на другой автомобиль без уведомления Заказчика.",
}
suites = [
    ("кабельные изделия",left_a,right_a,
     {("изменено","4.2","4.2"),("удалено","3.2",None),("добавлено",None,"5.2")},
     {("2.2","2.3")}),
    ("сборные перевозки",left_b,right_b,
     {("изменено","4.2","4.3"),("удалено","3.2",None),("добавлено",None,"5.2")},
     {("2.2","2.3")}),
]

class MemoryDocx(BytesIO):
    def __init__(self, name):
        super().__init__()
        self.name = name

    def resolve(self):
        return Path.cwd() / self.name


def make_clauses(name, clauses):
    document = Document()
    document.add_heading("Условия перевозки", 1)
    for clause_id in sorted(clauses, key=lambda item: tuple(map(int, item.split(".")))):
        document.add_paragraph(f"{clause_id}.\t{clauses[clause_id]}")
    with MemoryDocx(name) as stream:
        document.save(stream)
        stream.seek(0)
        return read_clauses(stream)


class ContractClauseTests(unittest.TestCase):
    def test_pdf_clause_and_appendix_sources_keep_page_numbers(self):
        blocks = list(pdf_blocks("contract.pdf", [
            (1, [("line", "1.1. Условие распространяется"), ("line", "на все перевозки.")]),
            (2, [("line", "Приложение № 1"), ("table", [["Маршрут", "Ставка"], ["Уфа", "72000"]])]),
        ]))
        self.assertEqual([item["kind"] for item in blocks], ["clause", "heading", "table_row", "table_row"])
        self.assertEqual(blocks[0]["text"], "1.1. Условие распространяется на все перевозки.")
        self.assertEqual(blocks[0]["page"], 1)
        self.assertEqual((blocks[-1]["page"], blocks[-1]["row_number"]), (2, 2))
        self.assertEqual(blocks[-1]["headings"], ["Приложение № 1"])

    def test_template_fit_requires_coverage_of_both_documents(self):
        template = [{"body": text} for text in ("A", "B", "C", "D", "E")]
        excerpt = template[:2]
        self.assertFalse(template_fit(template, excerpt, [[0.1] * 2 for _ in template])["suitable"])
        self.assertFalse(template_fit(excerpt, template, [[0.1] * 5 for _ in excerpt])["suitable"])
        self.assertTrue(template_fit(template, template[::-1], [[0.1] * 5 for _ in template])["suitable"])

    def test_template_fit_does_not_count_one_candidate_twice(self):
        template = [{"body": f"Left {i}"} for i in range(3)]
        carrier = [{"body": f"Right {i}"} for i in range(3)]
        fit = template_fit(template, carrier, [[0.95, 0.1, 0.1]] * 3)
        self.assertEqual(fit["matched_pairs"], 1)
        self.assertFalse(fit["suitable"])

    def test_additions_removals_changes_and_equivalent_wording(self):
        for name, left, right, expected, equivalent in suites:
            with self.subTest(name=name):
                template = make_clauses("template.docx", left)
                carrier = make_clauses("carrier.docx", right)
                changed = {(before, after) for state, before, after in expected if state == "изменено"}
                scores = [
                    [1.0 if a["body"] == b["body"] else
                     0.9 if (a["id"], b["id"]) in changed | equivalent else 0.1 for b in carrier]
                    for a in template
                ]

                def classify(batch):
                    result = {}
                    for case in batch:
                        source = case["source"]
                        left_side = source["source"]["source"].endswith("/template.docx")
                        side = 0 if left_side else 1
                        opposite = next((pair[1-side] for pair in changed | equivalent if pair[side] == source["id"]), None)
                        if opposite is None:
                            decision = {"status": "нет аналога", "source_quote": source["body"], "explanation": "Условие отсутствует"}
                        else:
                            candidate = next(item for item in case["candidates"] if item["id"] == opposite)
                            pair = (source["id"], opposite) if left_side else (opposite, source["id"])
                            decision = {"status": "совпадает", "candidate_id": opposite}
                            if pair in changed:
                                decision.update(status="изменено", source_quote=source["body"],
                                                candidate_quote=candidate["body"], explanation="Условие изменилось")
                        result[source["id"]] = decision
                    return result

                report = compare_clauses(template, carrier, scores, classify)
                actual = {
                    (finding["status"], finding.get("template", {}).get("clause_number"),
                     finding.get("carrier", {}).get("clause_number"))
                    for finding in report["findings"]
                }
                self.assertEqual(actual, expected)
                self.assertEqual(len(report["findings"]), len(expected))
                self.assertEqual(report["verbatim_pairs"], 6)
                self.assertEqual(report["equivalent_pairs"], 1)
                self.assertEqual(report["manual_review"], [])
                self.assertTrue(all(finding["explanation"] for finding in report["findings"]))

    def test_no_analogue_schema_excludes_candidate(self):
        source = "После выгрузки Перевозчик возвращает многооборотные поддоны."
        case = {
            "source": {"body": source},
            "candidates": [{"id": "3.2", "body": "Перевозчик разгружает автомобиль.", "already_paired_verbatim": True}],
        }
        decision = {"status": "нет аналога", "source_quote": source, "explanation": "Обязанность вернуть поддоны удалена"}
        options = decision_schema(case)["anyOf"]
        absence = next(branch for branch in options if branch["properties"]["status"]["enum"] == ["нет аналога"])
        self.assertNotIn("candidate_id", absence["properties"])
        self.assertFalse(absence["additionalProperties"])
        self.assertEqual(validate_decision(case, decision)[0], "нет аналога")
        self.assertEqual(validate_decision(case, {**decision, "candidate_id": "3.2"})[0], "неясно")

    def test_change_requires_a_verified_quote_and_explanation(self):
        case = {
            "source": {"body": "Срок — пять рабочих дней."},
            "candidates": [{"id": "4.2", "body": "Срок — пять календарных дней.", "already_paired_verbatim": False}],
        }
        decision = {
            "status": "изменено", "candidate_id": "4.2", "explanation": "Изменился тип дней",
            "source_quote": "пять рабочих дней", "candidate_quote": "пять календарных дней",
        }
        self.assertEqual(validate_decision(case, decision)[0], "изменено")
        for update in ({"candidate_quote": "семь дней"}, {"explanation": " "}):
            self.assertEqual(validate_decision(case, {**decision, **update})[0], "неясно")


if __name__ == "__main__":
    unittest.main()
