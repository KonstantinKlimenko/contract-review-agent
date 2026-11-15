"""Compare payment terms in two DOCX contracts using local models."""

import argparse
import importlib.util
import json
import os
import re
import sys
from collections import Counter
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen

from inspect_contract_docx import SAMPLE_CONTRACTS, extract_blocks


PAYMENT_QUERY = "В какой срок оплачивается выполненный рейс и от какого события идёт отсчёт?"
REFERENCE = re.compile(r"\bп\.\s*(\d+(?:\.\d+)+)")
EXTRACTION_PROMPT = (
    "Извлеки параметры оплаты основного выполненного рейса. Используй только текст "
    "предоставленных фрагментов. Верни число дней цифрами и тип дней. В start_events "
    "перечисли события, с которых по тексту начинается отсчёт срока оплаты. Для каждого "
    "события запиши действие существительным (например, подписание, получение или "
    "выставление) и отдельно предмет действия в именительном падеже. Сохраняй альтернативы "
    "отдельными элементами, если их несколько. При отсутствии события верни пустой массив. "
    "В quote каждого события скопируй точный фрагмент пункта, в котором названо событие. "
    "В source.id укажи ID пункта об оплате, в source.quote скопируй точные слова этого "
    "пункта о сроке. Сроки претензий и оформления документов пропусти."
)
ACTION_ALIASES = {"поступление": "получение"}
DOCUMENT_ALIASES = {
    "упд": "УПД",
    "универсальный передаточный документ": "УПД",
    "универсальный передаточный документ упд": "УПД",
    "универсального передаточного документа": "УПД",
    "счет": "счёт",
    "счет на оплату": "счёт",
    "счета": "счёт",
    "счета на оплату": "счёт",
}


def object_schema(properties):
    return {
        "type": "object",
        "properties": properties,
        "required": list(properties),
        "additionalProperties": False,
    }


STRING = {"type": "string"}
PAYMENT_SCHEMA = object_schema({
    "days": {"type": ["integer", "null"], "minimum": 0},
    "day_type": {"type": "string", "enum": ["календарные", "рабочие", ""]},
    "start_events": {
        "type": "array",
        "items": object_schema({"action": STRING, "object": STRING, "quote": STRING}),
    },
    "source": object_schema({"id": STRING, "quote": STRING}),
})


def searchable_text(block):
    return "\n".join([*block["headings"], block["text"]])


def select_candidates(blocks, scores):
    """Keep three search results and any unambiguous numbered references they cite."""
    counts = Counter(block["clause_number"] for block in blocks if block["kind"] == "clause")
    clauses = {block["clause_number"]: block for block in blocks
               if block["kind"] == "clause" and counts[block["clause_number"]] == 1}
    ranked = sorted(blocks, key=lambda block: scores[searchable_text(block)], reverse=True)
    candidates = {}
    for block in ranked[:3]:
        number = block.get("clause_number")
        candidate_id = number if number in clauses else f'{block["kind"]}@{block["body_position"]}'
        if block["kind"] == "table_row":
            candidate_id = f'{candidate_id}:{block["row_number"]}'
        candidates[candidate_id] = block
    for block in ranked[:3]:
        for number in REFERENCE.findall(block["text"]):
            if number in clauses:
                candidates.setdefault(number, clauses[number])
    return candidates


def extract_with_ollama(candidates, model):
    fragments = [{"id": key, "text": block["text"]} for key, block in candidates.items()]
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": EXTRACTION_PROMPT},
            {
                "role": "user",
                "content": (
                    "Фрагменты договора:\n"
                    + json.dumps(fragments, ensure_ascii=False)
                    + "\nИзвлеки срок оплаты и перечисли события начала отсчёта с цитатами."
                ),
            },
        ],
        "stream": False,
        "think": False,
        "format": PAYMENT_SCHEMA,
        "keep_alive": "2m",
        "options": {"temperature": 0, "seed": 42, "num_ctx": 4096, "num_predict": 650},
    }
    request = Request(
        "http://localhost:11434/api/chat",
        data=json.dumps(data).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    with urlopen(request, timeout=180) as response:
        reply = json.load(response)
    return json.loads(reply["message"]["content"])


def normalize(text):
    return " ".join(re.findall(r"[^\W_]+", text.casefold().replace("ё", "е")))


def read_terms(extraction, candidates):
    """Verify quote provenance and canonicalize the extracted event fields."""
    days = extraction["days"]
    day_type = extraction["day_type"]
    source = extraction["source"]
    block = candidates.get(source["id"], {})
    text = block.get("text", "")
    quote = source["quote"]
    verified = bool(quote.strip()) and quote in text
    events = []
    for event in extraction["start_events"]:
        action = normalize(event["action"])
        document = normalize(event["object"])
        event_quote = event["quote"]
        verified &= bool(action and document and event_quote.strip() and event_quote in text)
        events.append({
            "action": ACTION_ALIASES.get(action, action),
            "document": DOCUMENT_ALIASES.get(document, document),
            "quote": event_quote,
        })
    return {
        "days": days,
        "day_type": day_type,
        "start_events": events,
        "source": {
            "document": block.get("source"),
            "clause_number": block.get("clause_number"),
            "body_position": block.get("body_position"),
            **{key: block[key] for key in ("page", "page_end", "table_number", "row_number") if key in block},
            "quote": quote,
            "verified": verified,
        },
    }


def event_keys(terms):
    return {(event["action"], event["document"]) for event in terms["start_events"]}


def compare_terms(template, carrier):
    report = {"condition": "Оплата выполненного рейса", "template": template, "carrier": carrier}
    reasons = {}
    for side, terms in (("template", template), ("carrier", carrier)):
        if not terms["source"]["verified"]:
            reasons[side] = "Не удалось подтвердить цитаты по указанному источнику"
        elif type(terms["days"]) is not int or terms["days"] < 0 or not terms["day_type"]:
            reasons[side] = "Не удалось определить срок или тип дней"
        elif not terms["start_events"]:
            reasons[side] = "Не удалось определить событие начала отсчёта"
        elif len(event_keys(terms)) > 1:
            reasons[side] = "Найдено несколько возможных событий начала отсчёта"
    if reasons:
        return {**report, "status": "нужна ручная проверка", "reasons": reasons, "changes": []}

    template_event, = event_keys(template)
    carrier_event, = event_keys(carrier)
    fields = [
        ("срок в днях", template["days"], carrier["days"]),
        ("тип дней", template["day_type"], carrier["day_type"]),
        ("начало отсчёта", " ".join(template_event), " ".join(carrier_event)),
    ]
    changes = [{"field": name, "template": old, "carrier": new}
               for name, old, new in fields if old != new]
    return {**report, "status": "отличается" if changes else "совпадает", "changes": changes}


def review_payment(documents, scores, extract):
    terms = []
    for blocks in documents:
        candidates = select_candidates(blocks, scores)
        terms.append(read_terms(extract(candidates), candidates))
    return compare_terms(*terms)


def embedding_scores(documents, model_path):
    os.environ.setdefault("USE_TF", "0")
    os.environ.setdefault("HF_HUB_OFFLINE", "1")
    # The Windows RAG environment loads Arrow before Torch to avoid a DLL import error.
    if sys.platform == "win32" and importlib.util.find_spec("pyarrow"):
        import pyarrow
    import torch
    from sentence_transformers import SentenceTransformer

    torch.set_num_threads(4)
    encoder = SentenceTransformer(str(model_path))
    texts = list(dict.fromkeys(searchable_text(block) for blocks in documents for block in blocks))
    vectors = encoder.encode(
        [*texts, PAYMENT_QUERY], normalize_embeddings=True, batch_size=12, show_progress_bar=False,
    )
    if encoder.device.type == "cuda":
        encoder.to("cpu")
        torch.cuda.empty_cache()
    return dict(zip(texts, map(float, vectors[:-1] @ vectors[-1])))


def main():
    parser = argparse.ArgumentParser(description="Сравнивает условия оплаты в двух DOCX.")
    parser.add_argument("template", nargs="?", type=Path, default=Path(__file__).parent / "examples" / SAMPLE_CONTRACTS[0])
    parser.add_argument("carrier", nargs="?", type=Path, default=Path(__file__).parent / "examples" / SAMPLE_CONTRACTS[1])
    parser.add_argument("--embedding-model", type=Path, default=os.getenv("EMBEDDING_MODEL_PATH", Path(__file__).parent / "models" / "bge-m3"))
    parser.add_argument("--ollama-model", default="qwen3.5:4b")
    args = parser.parse_args()
    if not args.embedding_model.is_dir():
        parser.error(f"Не найдена папка модели эмбеддингов: {args.embedding_model}")

    documents = []
    for path in (args.template, args.carrier):
        if path.suffix.lower() != ".docx" or not path.is_file():
            parser.error(f"Нужен существующий файл DOCX: {path}")
        blocks = [block for block in extract_blocks(path) if block["kind"] != "heading"]
        if not blocks:
            parser.error(f"В документе нет текста для сравнения: {path}")
        documents.append(blocks)

    scores = embedding_scores(documents, args.embedding_model)
    try:
        report = review_payment(documents, scores, lambda candidates: extract_with_ollama(candidates, args.ollama_model))
    except (URLError, TimeoutError, json.JSONDecodeError) as error:
        parser.exit(1, f"Ошибка обращения к Ollama: {error}\n")
    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
