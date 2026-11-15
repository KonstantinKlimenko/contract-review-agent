import unittest
from pathlib import Path

from compare_contract_payment import review_payment, searchable_text
from inspect_contract_docx import SAMPLE_CONTRACTS, extract_blocks


EQUIVALENT_TEXT = (
    "Заказчик перечисляет плату за выполненную перевозку не позднее чем через "
    "тридцать календарных дней после того, как ему поступит универсальный передаточный документ."
)
AMBIGUOUS_TEXT = (
    "Заказчик оплачивает выполненный рейс в течение 30 календарных дней. "
    "Срок исчисляется либо с даты выставления счёта, либо со дня получения УПД. "
    "Применимую дату стороны определяют отдельно для каждого рейса."
)


def event(action, document, quote):
    return {"action": action, "object": document, "quote": quote}


def payment(number, days, quote, *events):
    return {
        "days": days,
        "day_type": "календарные",
        "source": {"id": number, "quote": quote},
        "start_events": list(events),
    }


class PaymentComparisonTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template, cls.carrier = [
            [block for block in extract_blocks(Path(__file__).parent / "examples" / name)
             if block["kind"] != "heading"]
            for name in SAMPLE_CONTRACTS
        ]

    def run_review(self, carrier_payment, carrier_text=None):
        carrier = self.carrier
        if carrier_text:
            carrier = [
                {**block, "text": "4.2.\t" + carrier_text}
                if block.get("clause_number") == "4.2" else block
                for block in carrier
            ]
        scores = {
            searchable_text(block): int(block.get("clause_number") == clause)
            for blocks, clause in ((self.template, "5.2"), (carrier, "4.2"))
            for block in blocks
        }
        template_payment = payment(
            "5.2", 30, "в течение 30 календарных дней с даты получения УПД",
            event("получение", "УПД", "с даты получения УПД"),
        )
        model_responses = iter((template_payment, carrier_payment))
        return review_payment((self.template, carrier), scores, lambda _: next(model_responses))

    def test_changed_duration_and_start_event(self):
        report = self.run_review(payment(
            "4.2", 10, "не позднее 10 календарных дней",
            event("выставление", "счёт на оплату", "с даты выставления Перевозчиком счёта на оплату"),
        ))
        self.assertEqual(report["status"], "отличается")
        self.assertEqual(report["changes"], [
            {"field": "срок в днях", "template": 30, "carrier": 10},
            {"field": "начало отсчёта", "template": "получение УПД", "carrier": "выставление счёт"},
        ])
        self.assertEqual(report["template"]["source"]["clause_number"], "5.2")
        self.assertEqual(report["carrier"]["source"]["clause_number"], "4.2")

    def test_equivalent_wording(self):
        report = self.run_review(payment(
            "4.2", 30, "через тридцать календарных дней",
            event("поступление", "универсальный передаточный документ",
                  "после того, как ему поступит универсальный передаточный документ"),
        ), EQUIVALENT_TEXT)
        self.assertEqual(report["status"], "совпадает")
        self.assertEqual(report["changes"], [])

    def test_two_possible_start_events(self):
        report = self.run_review(payment(
            "4.2", 30, "в течение 30 календарных дней",
            event("выставление", "счёт", "с даты выставления счёта"),
            event("получение", "УПД", "со дня получения УПД"),
        ), AMBIGUOUS_TEXT)
        self.assertEqual(report["status"], "нужна ручная проверка")
        self.assertEqual(report["changes"], [])
        self.assertIn("несколько", report["reasons"]["carrier"])
        self.assertEqual(len(report["carrier"]["start_events"]), 2)
        self.assertTrue(report["carrier"]["source"]["verified"])

    def test_unverified_quote_requires_manual_review(self):
        report = self.run_review(payment(
            "4.2", 10, "не позднее 10 календарных дней",
            event("выставление", "счёт", "с момента подписания счёта"),
        ))
        self.assertEqual(report["status"], "нужна ручная проверка")
        self.assertFalse(report["carrier"]["source"]["verified"])
        self.assertEqual(report["changes"], [])


if __name__ == "__main__":
    unittest.main()
