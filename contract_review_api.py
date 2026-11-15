"""Local API for the contract review forms in n8n."""

from base64 import b64decode
from binascii import Error as Base64Error
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
from decimal import Decimal, InvalidOperation
import hashlib
from html import escape
from io import BytesIO
import json
import logging
import os
from pathlib import Path
import re
from secrets import token_urlsafe
import sqlite3
from typing import Literal
from urllib.error import URLError
from urllib.request import Request as UpstreamRequest, urlopen
from uuid import UUID, uuid4
from zipfile import BadZipFile, ZipFile

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse
from docx.opc.exceptions import PackageNotFoundError
from lxml.etree import XMLSyntaxError
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from pdfminer.pdfexceptions import PDFException
from pydantic import BaseModel, Field

import compare_contract_clauses as clauses
import compare_contract_payment as payment
from compare_contract_details import review_details
from inspect_contract_docx import appendix_headings, extract_blocks, missing_appendices


ROOT = Path(os.getenv("CONTRACT_REVIEW_DATA", Path(__file__).parent / "data"))
MODEL = Path(os.getenv("EMBEDDING_MODEL_PATH", Path(__file__).parent / "models" / "bge-m3"))
OLLAMA_MODEL = os.getenv("CONTRACT_REVIEW_MODEL", "qwen3.5:4b")
MAX_UPLOAD = 10 * 1024 * 1024
MAX_RATES_FILE = 2 * 1024 * 1024
RATES_COLUMNS = ("Откуда", "Куда", "Тип машины", "Единица расчёта", "Тариф", "Валюта", "НДС")
app = FastAPI()
worker = ThreadPoolExecutor(max_workers=1)
logger = logging.getLogger(__name__)


class UploadedContract(BytesIO):
    def __init__(self, content, name):
        super().__init__(content)
        self.name = name

    def resolve(self):
        return Path(self.name)


def db():
    ROOT.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(ROOT / "reviews.sqlite")
    connection.row_factory = sqlite3.Row
    connection.executescript("""
        CREATE TABLE IF NOT EXISTS templates
            (version TEXT PRIMARY KEY, name TEXT NOT NULL, content BLOB NOT NULL);
        CREATE TABLE IF NOT EXISTS active
            (slot INTEGER PRIMARY KEY CHECK (slot = 1), version TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS jobs
            (id TEXT PRIMARY KEY, state TEXT NOT NULL, result TEXT);
        CREATE TABLE IF NOT EXISTS chat_reviews
            (session_id TEXT PRIMARY KEY, review_id TEXT NOT NULL, contract_name TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS chat_bridge
            (slot INTEGER PRIMARY KEY CHECK (slot = 1), token TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS review_decisions
            (id INTEGER PRIMARY KEY, review_id TEXT NOT NULL, finding_id TEXT NOT NULL,
             decision TEXT NOT NULL, comment TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP);
        CREATE INDEX IF NOT EXISTS review_decisions_lookup ON review_decisions(review_id, finding_id, id);
        CREATE TABLE IF NOT EXISTS review_drafts
            (review_id TEXT PRIMARY KEY, decision_ids TEXT NOT NULL, body TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS review_refusals
            (review_id TEXT NOT NULL, finding_id TEXT NOT NULL, PRIMARY KEY(review_id, finding_id));
        CREATE TABLE IF NOT EXISTS review_batches
            (id TEXT PRIMARY KEY, template_version TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS review_batch_items
            (batch_id TEXT NOT NULL, position INTEGER NOT NULL, name TEXT NOT NULL,
             review_id TEXT, error TEXT, PRIMARY KEY(batch_id, position));
        CREATE TABLE IF NOT EXISTS review_batch_rates
            (batch_id TEXT PRIMARY KEY, result TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS rate_templates
            (version TEXT PRIMARY KEY, name TEXT NOT NULL, content TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS active_rate_template
            (slot INTEGER PRIMARY KEY CHECK (slot = 1), version TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS batch_rate_templates
            (batch_id TEXT PRIMARY KEY, version TEXT NOT NULL);
    """)
    return connection


@app.on_event("startup")
def mark_interrupted_jobs():
    with closing(db()) as connection, connection:
        connection.execute("UPDATE jobs SET state='failed' WHERE state IN ('queued','running')")


async def uploaded(request, name):
    content = bytearray()
    async for chunk in request.stream():
        content.extend(chunk)
        if len(content) > MAX_UPLOAD:
            raise HTTPException(413, "Допустимый размер договора — 10 МБ")
    return validate_contract(bytes(content), name)


def validate_contract(data, name):
    suffix = Path(name).suffix.lower()
    if suffix == ".docx":
        return validate_docx(data)
    if suffix != ".pdf":
        raise HTTPException(422, "Нужен DOCX или PDF с текстовым слоем.")
    if len(data) > MAX_UPLOAD:
        raise HTTPException(413, "Допустимый размер PDF — 10 МБ.")
    try:
        blocks = list(extract_blocks(UploadedContract(data, name)))
    except ValueError as error:
        raise HTTPException(422, str(error)) from error
    except (PDFException, OSError, TypeError, KeyError) as error:
        raise HTTPException(422, "Не удалось прочитать PDF. Проверьте формат и защиту файла.") from error
    if not blocks:
        raise HTTPException(422, "В PDF нет извлекаемого текста. Нужен PDF с текстовым слоем.")
    if not any(block["kind"] == "clause" for block in blocks):
        raise HTTPException(422, "В PDF не найдены нумерованные пункты договора, например 1.1.")
    return data


def validate_docx(data):
    if len(data) > MAX_UPLOAD:
        raise HTTPException(413, "Допустимый размер DOCX — 10 МБ")
    try:
        with ZipFile(BytesIO(data)) as package:
            entries = package.infolist()
            if len(entries) > 2000 or sum(entry.file_size for entry in entries) > 50 * 1024 * 1024:
                raise HTTPException(413, "Распакованный DOCX слишком велик")
            if "word/document.xml" not in package.namelist():
                raise HTTPException(422, "Нужен файл DOCX")
        if not clauses.read_clauses(UploadedContract(data, "upload.docx")):
            raise HTTPException(422, "В документе нет распознаваемых нумерованных пунктов")
    except (BadZipFile, PackageNotFoundError, XMLSyntaxError, KeyError, ValueError) as error:
        raise HTTPException(422, "Не удалось прочитать DOCX") from error
    return data


def active_template():
    with closing(db()) as connection:
        return active_word(connection)


def active_word(connection):
    return connection.execute(
        "SELECT t.* FROM active a JOIN templates t USING(version) WHERE a.slot=1"
    ).fetchone()


@app.get("/api/template")
def template_status():
    row = active_template()
    return {"active": {"name": row["name"], "version": row["version"]} if row else None}


def word_templates(connection):
    return connection.execute(
        "SELECT t.name,t.version,a.version IS NOT NULL AS selected FROM templates t LEFT JOIN active a "
        "ON a.version=t.version ORDER BY selected DESC,t.name,t.version"
    ).fetchall()


@app.get("/api/templates")
def template_catalog():
    with closing(db()) as connection:
        rows = word_templates(connection)
    return {"templates": [
        {"name": row["name"], "version": row["version"], "active": bool(row["selected"]),
         "label": f'{row["name"]} · {row["version"][:12]}' + (" · активный" if row["selected"] else "")}
        for row in rows
    ]}


@app.post("/api/template")
async def set_template(request: Request):
    name = request.query_params.get("name", "Шаблон.docx")[:200]
    content = await uploaded(request, name)
    with closing(db()) as connection, connection:
        version = activate_word(connection, name, content)
    return {"version": version, "redirect": "http://localhost:8765/template"}


def activate_word(connection, name, content):
    version = hashlib.sha256(content).hexdigest()
    connection.execute("INSERT OR IGNORE INTO templates VALUES (?,?,?)", (version, name, content))
    connection.execute(
        "INSERT INTO active VALUES (1,?) ON CONFLICT(slot) DO UPDATE SET version=excluded.version",
        (version,),
    )
    return version


def run_checks(template, contract):
    paths = (template, contract)
    numbered = [clauses.read_clauses(path) for path in paths]
    similarity = clauses.embedding_similarity(*numbered, MODEL)
    fit = clauses.template_fit(*numbered, similarity)
    if not fit["suitable"]:
        return {"template_fit": fit}
    blocks = [list(extract_blocks(path)) for path in paths]
    search_blocks = [[block for block in doc if block["kind"] != "heading"] for doc in blocks]
    clause_report = clauses.compare_clauses(
        *numbered, similarity, lambda batch: clauses.classify_with_ollama(batch, OLLAMA_MODEL)
    )
    scores = payment.embedding_scores(search_blocks, MODEL)
    payment_report = payment.review_payment(
        search_blocks, scores, lambda batch: payment.extract_with_ollama(batch, OLLAMA_MODEL)
    )
    pdf_tables = {}
    for block in blocks[1]:
        if block["kind"] == "table_row" and "page" in block:
            table = pdf_tables.setdefault(block["table_number"], {
                "number": block["table_number"], "page": block["page"], "rows": 0,
            })
            table["rows"] += 1
    return {
        "template_fit": fit,
        "clauses": clause_report,
        "payment": payment_report,
        "party_details": review_details(*blocks),
        "missing_appendices": missing_appendices(*blocks),
        "appendix_titles": [title for _, title in appendix_headings(blocks[1])],
        "appendix_sources": [{"title": block["text"], "page": block["page"]} for block in blocks[1]
                             if "page" in block and list(appendix_headings([block]))],
        "pdf_tables": list(pdf_tables.values()),
        "clause_fingerprints": {
            side: {str(item["source"]["body_position"]): fingerprint(item["body"]) for item in items}
            for side, items in zip(("template", "carrier"), numbered)
        },
        "document_structure": {
            side: {
                "tables": len({block["table_number"] for block in doc if block["kind"] == "table_row"}),
                "appendices": len(list(appendix_headings(doc))),
                "unnumbered_paragraphs": sum(block["kind"] == "paragraph" for block in doc),
            }
            for side, doc in zip(("template", "carrier"), blocks)
        },
    }


def template_fit_message(fit):
    pairs = fit["matched_pairs"]
    return (f"Найдено близких пар: {pairs} из {fit['template_count']} пунктов шаблона "
            f"и {pairs} из {fit['carrier_count']} пунктов договора. "
            "Для детальной сверки выберите подходящий шаблон и загрузите договор снова.")


def update_job(job_id, state, result=None):
    with closing(db()) as connection, connection:
        connection.execute("UPDATE jobs SET state=?, result=? WHERE id=?",
                           (state, json.dumps(result, ensure_ascii=False) if result else None, job_id))


def process_job(job_id, template, content, name, previous=None):
    update_job(job_id, "running")
    try:
        result = run_checks(UploadedContract(template["content"], template["name"]), UploadedContract(content, name))
        result["template_version"] = template["version"]
        result["template_name"] = template["name"]
        result["contract_name"] = name
        if not result["template_fit"]["suitable"]:
            update_job(job_id, "template_mismatch", result)
            return
        if previous:
            result["revision"] = compare_revision(result, *previous)
        with closing(db()) as connection, connection:
            if previous:
                for item in result["revision"]["items"]:
                    if item["status"] != "unchanged":
                        continue
                    if saved := item["previous_decision"]:
                        connection.execute(
                            "INSERT INTO review_decisions(review_id,finding_id,decision,comment) VALUES (?,?,?,?)",
                            (job_id, item["current_id"], saved["decision"], saved["comment"]),
                        )
                    if item["previous_refusal"]:
                        connection.execute("INSERT INTO review_refusals VALUES (?,?)", (job_id, item["current_id"]))
            connection.execute("UPDATE jobs SET state='completed', result=? WHERE id=?",
                               (json.dumps(result, ensure_ascii=False), job_id))
    except Exception:
        logger.exception("Review job %s failed", job_id)
        update_job(job_id, "failed")


@app.post("/api/reviews", status_code=202)
async def submit_review(
    request: Request,
    session_id: str | None = Query(default=None, min_length=1, max_length=200),
):
    template = active_template()
    if template is None:
        raise HTTPException(409, "Сначала загрузите шаблон заказчика")
    name = request.query_params.get("name", "Договор.docx")[:200]
    content = await uploaded(request, name)
    job_id = uuid4().hex
    with closing(db()) as connection, connection:
        connection.execute("INSERT INTO jobs(id,state) VALUES (?, 'queued')", (job_id,))
        if session_id is not None:
            connection.execute(
                """INSERT INTO chat_reviews (session_id, review_id, contract_name) VALUES (?, ?, ?)
                   ON CONFLICT(session_id) DO UPDATE SET
                       review_id=excluded.review_id, contract_name=excluded.contract_name""",
                (session_id, job_id, name),
            )
    worker.submit(process_job, job_id, dict(template), content, name)
    return {"id": job_id, "redirect": f"http://localhost:8765/reviews/{job_id}"}


@app.get("/api/chat-review")
def chat_review(
    session_id: str = Query(min_length=1, max_length=200),
    section: Literal["payment", "details", "clauses", "summary"] = "payment",
):
    with closing(db()) as connection:
        row = connection.execute(
            """SELECT j.id, j.state, j.result, c.contract_name
               FROM chat_reviews c JOIN jobs j ON j.id=c.review_id
               WHERE c.session_id=?""",
            (session_id,),
        ).fetchone()
    if row is None:
        return {"state": "missing"}
    result = json.loads(row["result"]) if row["result"] else {}
    response = {
        "id": row["id"],
        "state": row["state"],
        "contract_name": row["contract_name"],
        "redirect": f"http://localhost:8765/reviews/{row['id']}",
    }
    if row["state"] == "template_mismatch":
        return {**response, "template_name": result["template_name"], "template_fit": result["template_fit"]}
    if section == "summary":
        if row["state"] == "completed":
            report = result.get("clauses", {})
            response["finding_count"] = len(report.get("findings", []))
            response["manual_clause_count"] = len(report.get("manual_review", []))
            response["party_issues"] = [
                {"role": item["role"], "field": item["field"], "status": item["status"]}
                for item in result.get("party_details", []) if item["status"] != "совпадает"
            ]
            if "missing_appendices" in result:
                response["missing_appendices"] = result["missing_appendices"]
    elif section == "details":
        response["party_details"] = [
            {
                **{key: item[key] for key in ("role", "field", "status", "reasons")},
                **{
                    side: list(dict.fromkeys(entry["value"] for entry in item[side]))
                    for side in ("template", "carrier")
                },
            }
            for item in result.get("party_details", [])
        ]
    elif section == "clauses":
        report = result.get("clauses", {})
        response["findings"] = report.get("findings", [])
        response["manual_review"] = report.get("manual_review", [])
        if "clause_matches" in report:
            response["clause_matches"] = report["clause_matches"]
        if "missing_appendices" in result:
            response["missing_appendices"] = result["missing_appendices"]
        if result.get("appendix_sources"):
            response["appendix_sources"] = result["appendix_sources"]
    else:
        response["payment"] = payment_evidence(result.get("payment"))
    return response


@app.get("/api/reviews/{job_id}")
def review_status(job_id: str):
    with closing(db()) as connection:
        row = connection.execute(
            """SELECT state, result,
                      (SELECT contract_name FROM chat_reviews WHERE review_id=jobs.id LIMIT 1) AS contract_name
               FROM jobs WHERE id=?""",
            (job_id,),
        ).fetchone()
    if row is None:
        raise HTTPException(404, "Проверка не найдена")
    result = json.loads(row["result"]) if row["result"] else None
    if result is not None and not result.get("contract_name"):
        sources = {
            record["source"]["document"]
            for item in result.get("party_details", []) for record in item["carrier"]
            if record["source"].get("document")
        }
        result["contract_name"] = row["contract_name"] or (sources.pop() if len(sources) == 1 else None)
    return {"state": row["state"], "result": result}


def chat_bridge_token():
    with closing(db()) as connection, connection:
        connection.execute("INSERT OR IGNORE INTO chat_bridge VALUES (1,?)", (token_urlsafe(32),))
        return connection.execute("SELECT token FROM chat_bridge WHERE slot=1").fetchone()["token"]


class ReviewQuestion(BaseModel):
    session_id: str = Field(min_length=1, max_length=200)
    message: str = Field(min_length=1, max_length=2000)


class ReviewSession(BaseModel):
    session_id: str = Field(min_length=1, max_length=200)


class BatchRequest(BaseModel):
    files: list[str] = Field(min_length=1, max_length=10)
    template_version: str | None = None


class BatchQuestion(BaseModel):
    message: str = Field(min_length=1, max_length=2000)
    session_id: UUID | None = None


class RatesFile(BaseModel):
    name: str = Field(min_length=1, max_length=255)
    content: str = Field(max_length=4 * ((MAX_RATES_FILE + 2) // 3))


class WordTemplateFile(BaseModel):
    name: str = Field(min_length=1, max_length=255)
    content: str = Field(max_length=4 * ((MAX_UPLOAD + 2) // 3))


class TemplateFiles(BaseModel):
    word: WordTemplateFile | None = None
    rates: RatesFile | None = None


class BatchRatesUpload(BaseModel):
    approved: RatesFile | None = None
    vendor: RatesFile | None = None


def rates_key(value):
    return " ".join(str(value).casefold().replace("ё", "е").split())


def read_rates(file):
    data = b64decode(file.content, validate=True)
    if not file.name.lower().endswith(".xlsx") or len(data) > MAX_RATES_FILE:
        raise ValueError("Нужен XLSX размером до 2 МиБ.")
    with ZipFile(BytesIO(data)) as archive:
        entries = archive.infolist()
        if len(entries) > 500 or sum(entry.file_size for entry in entries) > 20 * 1024 * 1024:
            raise ValueError("Распакованная книга слишком велика.")
    with closing(load_workbook(BytesIO(data), read_only=True, keep_links=False)) as book:
        if "Тарифы" not in book:
            raise ValueError("Добавьте лист «Тарифы».")
        sheet = book["Тарифы"]
        sheet.reset_dimensions()
        rows = sheet.iter_rows(max_col=40)
        headers = [rates_key(cell.value) for cell in next(rows, ())]
        expected = [rates_key(name) for name in RATES_COLUMNS]
        if any(headers.count(name) != 1 for name in expected):
            raise ValueError("В первой строке нужны уникальные колонки " + ", ".join(RATES_COLUMNS) + ".")
        indexes = [headers.index(name) for name in expected]
        result = {}
        for number, row in enumerate(rows, start=2):
            if number > 501:
                raise ValueError("Допустимо до 500 строк тарифов под заголовком.")
            cells = [row[index] for index in indexes]
            if all(cell.value is None for cell in cells):
                continue
            if any(cell.value is None or cell.data_type in ("f", "e") for cell in cells):
                raise ValueError(f"Строка {number} содержит пустые ячейки, ошибки или формулы.")
            origin, destination, vehicle, unit, value, currency, vat = (cell.value for cell in cells)
            labels = [origin, destination, vehicle, unit, currency, vat]
            if any(not isinstance(label, str) or not label.strip() or len(label) > 100 for label in labels):
                raise ValueError(f"Строка {number} содержит некорректные текстовые поля.")
            amount = Decimal(str(value).replace(" ", "").replace("\xa0", "").replace(",", "."))
            if not amount.is_finite() or amount < 0 or amount > 10**12 or amount.as_tuple().exponent < -4:
                raise ValueError(f"Строка {number} содержит недопустимый тариф.")
            key = tuple(map(rates_key, (origin, destination, vehicle)))
            if key in result:
                raise ValueError(f"Строка {number} повторяет маршрут и тип машины.")
            result[key] = dict(origin=origin.strip(), destination=destination.strip(), vehicle=vehicle.strip(),
                               unit=unit.strip(), rate=format(amount, "f"), currency=currency.strip(), vat=vat.strip(),
                               file=file.name, sheet=sheet.title, cell=f"{get_column_letter(indexes[4] + 1)}{number}")
        if not result:
            raise ValueError("На листе нет ставок.")
        return result


def compare_rates(approved, vendor):
    if not approved or not vendor:
        return {"state": "incomplete", "message": "Для сверки нужны обе тарифные таблицы.", "findings": []}
    tables = []
    for file in (approved, vendor):
        try:
            tables.append(read_rates(file))
        except (Base64Error, BadZipFile, InvalidFileException, InvalidOperation, XMLSyntaxError, KeyError, ValueError) as error:
            return {"state": "invalid", "message": f"{file.name}: {rate_file_error(error)}", "findings": []}
    reference, proposed = tables
    findings = []
    matched = 0
    for key in sorted(reference.keys() | proposed.keys()):
        left, right = reference.get(key), proposed.get(key)
        difference = None
        if left is None or right is None:
            status = "added" if left is None else "missing"
        elif any(rates_key(left[field]) != rates_key(right[field]) for field in ("currency", "unit", "vat")):
            status = "incompatible"
        else:
            delta = Decimal(right["rate"]) - Decimal(left["rate"])
            if not delta:
                matched += 1
                continue
            status, difference = "changed", format(delta, "f")
        findings.append(dict(status=status, approved=left, vendor=right, difference=difference))
    return dict(state="completed", approved_file=approved.name, vendor_file=vendor.name,
                matched=matched, findings=findings)


def rate_file_error(error):
    return str(error) if type(error) is ValueError else "Не удалось прочитать тарифы из XLSX."


def rate_template_metadata(row):
    return {"version": row["version"], "name": row["name"]} if row else None


def active_rates(connection):
    return connection.execute(
        "SELECT r.name,r.version FROM active_rate_template a JOIN rate_templates r USING(version) WHERE a.slot=1"
    ).fetchone()


@app.get("/api/rate-template")
def rate_template_status():
    with closing(db()) as connection:
        return rate_template_metadata(active_rates(connection))


@app.post("/api/rate-template", status_code=201)
def set_rate_template(file: RatesFile):
    version = rate_template_version(file)
    with closing(db()) as connection, connection:
        activate_rates(connection, file, version)
        return rate_template_metadata(active_rates(connection))


def rate_template_version(file):
    try:
        read_rates(file)
    except (Base64Error, BadZipFile, InvalidFileException, InvalidOperation, XMLSyntaxError, KeyError, ValueError) as error:
        raise HTTPException(422, rate_file_error(error)) from error
    return hashlib.sha256(b64decode(file.content)).hexdigest()


def activate_rates(connection, file, version):
    connection.execute("INSERT OR IGNORE INTO rate_templates VALUES (?,?,?)", (version, file.name, file.content))
    connection.execute(
        "INSERT INTO active_rate_template VALUES (1,?) ON CONFLICT(slot) DO UPDATE SET version=excluded.version",
        (version,),
    )


@app.post("/api/templates", status_code=201)
def set_templates(files: TemplateFiles):
    if not files.word and not files.rates:
        raise HTTPException(422, "Приложите шаблон DOCX или PDF, таблицу XLSX либо оба файла.")
    document = None
    if files.word:
        try:
            document = validate_contract(b64decode(files.word.content, validate=True), files.word.name)
        except Base64Error as error:
            raise HTTPException(422, "Не удалось прочитать файл шаблона.") from error
    rates_version = rate_template_version(files.rates) if files.rates else None
    with closing(db()) as connection, connection:
        if document is not None:
            activate_word(connection, files.word.name, document)
        if rates_version:
            activate_rates(connection, files.rates, rates_version)
        return {"word": rate_template_metadata(active_word(connection)),
                "rates": rate_template_metadata(active_rates(connection)),
                "redirect": "http://localhost:8765/template"}


@app.post("/api/batches", status_code=201)
def create_batch(request: BatchRequest):
    batch_id = uuid4().hex
    with closing(db()) as connection, connection:
        connection.execute("BEGIN IMMEDIATE")
        template = (connection.execute("SELECT version FROM templates WHERE version=?", (request.template_version,)).fetchone()
                    if request.template_version else active_word(connection))
        if template is None:
            raise HTTPException(409, "Выберите доступный шаблон.")
        connection.execute("INSERT INTO review_batches VALUES (?,?)", (batch_id, template["version"]))
        rate_template = active_rates(connection)
        if rate_template:
            connection.execute("INSERT INTO batch_rate_templates VALUES (?,?)", (batch_id, rate_template["version"]))
        connection.executemany("INSERT INTO review_batch_items VALUES (?,?,?,NULL,NULL)", [
            (batch_id, index, name.strip()[:200] or "Договор.docx") for index, name in enumerate(request.files)
        ])
    return {"id": batch_id}


@app.put("/api/batches/{batch_id}/files/{position}", status_code=202)
async def upload_batch_file(batch_id: str, position: int, request: Request):
    with closing(db()) as connection:
        item = connection.execute(
            "SELECT i.*, b.template_version FROM review_batch_items i JOIN review_batches b ON b.id=i.batch_id "
            "WHERE i.batch_id=? AND i.position=?", (batch_id, position),
        ).fetchone()
    if item is None:
        raise HTTPException(404, "Файл пакета не найден.")
    if item["review_id"]:
        return {"id": item["review_id"]}
    try:
        content = await uploaded(request, item["name"])
    except HTTPException as error:
        with closing(db()) as connection, connection:
            connection.execute("UPDATE review_batch_items SET error=? WHERE batch_id=? AND position=? AND review_id IS NULL",
                               (error.detail, batch_id, position))
        raise
    job_id, session_id = uuid4().hex, str(uuid4())
    with closing(db()) as connection, connection:
        connection.execute("BEGIN IMMEDIATE")
        existing = connection.execute("SELECT review_id FROM review_batch_items WHERE batch_id=? AND position=?",
                                      (batch_id, position)).fetchone()["review_id"]
        if existing:
            return {"id": existing}
        template = connection.execute("SELECT * FROM templates WHERE version=?", (item["template_version"],)).fetchone()
        connection.execute("INSERT INTO jobs(id,state) VALUES (?, 'queued')", (job_id,))
        connection.execute("INSERT INTO chat_reviews VALUES (?,?,?)", (session_id, job_id, item["name"]))
        connection.execute("UPDATE review_batch_items SET review_id=?, error=NULL WHERE batch_id=? AND position=?",
                           (job_id, batch_id, position))
    worker.submit(process_job, job_id, dict(template), content, item["name"])
    return {"id": job_id}


def batch_data(batch_id):
    with closing(db()) as connection:
        batch = connection.execute("SELECT t.name,t.version FROM review_batches b JOIN templates t "
                                   "ON b.template_version=t.version WHERE b.id=?", (batch_id,)).fetchone()
        if batch is None:
            raise HTTPException(404, "Пакет не найден.")
        rows = connection.execute(
            "SELECT i.*,j.state,j.result,c.session_id FROM review_batch_items i LEFT JOIN jobs j ON j.id=i.review_id "
            "LEFT JOIN chat_reviews c ON c.review_id=j.id WHERE i.batch_id=? ORDER BY i.position", (batch_id,),
        ).fetchall()
    return batch, rows


def batch_rates(batch_id):
    with closing(db()) as connection:
        row = connection.execute("SELECT result FROM review_batch_rates WHERE batch_id=?", (batch_id,)).fetchone()
    return json.loads(row["result"]) if row else {"state": "not_uploaded", "findings": []}


def batch_rate_template(batch_id, *, include_content=False):
    columns = "r.name,r.version,r.content" if include_content else "r.name,r.version"
    with closing(db()) as connection:
        return connection.execute(
            f"SELECT {columns} FROM batch_rate_templates b JOIN rate_templates r USING(version) WHERE b.batch_id=?",
            (batch_id,),
        ).fetchone()


@app.put("/api/batches/{batch_id}/rates")
def upload_batch_rates(batch_id: str, request: BatchRatesUpload):
    batch_data(batch_id)
    approved = request.approved
    pinned_version = None
    if approved is None:
        pinned = batch_rate_template(batch_id, include_content=True)
        if pinned:
            approved = RatesFile(name=pinned["name"], content=pinned["content"])
            pinned_version = pinned["version"]
    result = compare_rates(approved, request.vendor)
    if result["state"] == "completed" and pinned_version:
        result["approved_version"] = pinned_version
    with closing(db()) as connection, connection:
        connection.execute(
            "INSERT INTO review_batch_rates VALUES (?,?) "
            "ON CONFLICT(batch_id) DO UPDATE SET result=excluded.result",
            (batch_id, json.dumps(result, ensure_ascii=False)),
        )
    return {"state": result["state"]}


@app.get("/api/batches/{batch_id}/rates")
def batch_rates_evidence(batch_id: str, route: str = Query(default="", max_length=120)):
    batch_data(batch_id)
    result = batch_rates(batch_id)
    terms = re.findall(r"\w+", rates_key(route))
    findings = []
    for item in result["findings"]:
        source = item["vendor"] or item["approved"]
        if all(term in rates_key(source["origin"] + " " + source["destination"]) for term in terms):
            findings.append(item)
    return {**result, "findings": findings[:20], "total_findings": len(result["findings"]),
            "matching_findings": len(findings), "truncated": len(findings) > 20,
            "report": f"http://localhost:8765/batch?id={batch_id}#rates"}


@app.get("/api/batches/{batch_id}")
def batch_status(batch_id: str):
    batch, rows = batch_data(batch_id)
    items = []
    for row in rows:
        result = json.loads(row["result"]) if row["result"] else {}
        items.append({
            "name": row["name"], "state": row["state"] or ("failed" if row["error"] else "waiting"),
            "error": row["error"], "review_id": row["review_id"], "session_id": row["session_id"],
            "findings": list(review_findings(result).values()) if row["state"] == "completed" else [],
            "manual": len(result["clauses"]["manual_review"]) if row["state"] == "completed" else 0,
            "payment": (result.get("payment") or {}).get("status"),
            "template_fit": result.get("template_fit") if row["state"] == "template_mismatch" else None,
        })
    return {"template": dict(batch), "items": items, "rates": batch_rates(batch_id),
            "rate_template": rate_template_metadata(batch_rate_template(batch_id))}


def payment_evidence(terms):
    if not terms:
        return terms
    evidence = dict(terms)
    for side in ("template", "carrier"):
        if detail := terms.get(side):
            source = {key: value for key, value in detail["source"].items() if key != "body_position"}
            evidence[side] = {**detail, "source": source}
    return evidence


@app.get("/api/batches/{batch_id}/evidence")
def batch_evidence(batch_id: str, section: Literal["summary", "payment", "appendices", "findings"] = "summary",
                   file_number: int | None = Query(default=None, ge=1, le=10)):
    batch, rows = batch_data(batch_id)
    items = []
    for row in rows:
        if file_number is not None and row["position"] + 1 != file_number:
            continue
        result = json.loads(row["result"]) if row["result"] else {}
        items.append({
            "number": row["position"] + 1, "name": row["name"],
            "state": row["state"] or ("failed" if row["error"] else "waiting"),
            "report": f'http://localhost:8765/reviews/{row["review_id"]}' if result else None,
        })
        if row["state"] == "template_mismatch":
            items[-1]["template_fit"] = result["template_fit"]
        elif row["state"] == "completed":
            findings = review_findings(result)
            items[-1].update(finding_count=len(findings),
                             missing_appendices_count=len(result.get("missing_appendices", [])))
            if section == "payment":
                terms = payment_evidence(result.get("payment"))
                if terms:
                    items[-1]["payment"] = {
                        **{key: value for key, value in terms.items() if key != "status"},
                        "payment_comparison": terms["status"],
                    }
            elif section == "summary":
                items[-1]["payment_comparison"] = (result.get("payment") or {}).get("status")
            elif section == "appendices":
                items[-1]["appendices"] = result.get("missing_appendices", [])
                items[-1]["appendix_sources"] = result.get("appendix_sources", [])
            elif section == "findings":
                items[-1]["findings"] = list(findings.values()) if file_number else [x["title"] for x in findings.values()]
    return {"template": batch["name"], "items": items}


@app.post("/api/batches/{batch_id}/chat")
def ask_batch(batch_id: str, question: BatchQuestion):
    batch_data(batch_id)
    return {"output": call_review_workflow(
        {
            "action": "batch",
            "batchId": batch_id,
            "chatInput": question.message,
            "sessionId": str(question.session_id) if question.session_id else None,
        },
        "Агент пакета недоступен. Попробуйте ещё раз.",
    )}


class FindingDecision(ReviewSession):
    decision: Literal["accept", "propose", "discuss"]
    comment: str = Field(default="", max_length=2000)


class CarrierRefusal(ReviewSession):
    refused: bool


def with_pdf_page(location, source):
    if "page" not in source:
        return location
    start, end = source["page"], source.get("page_end")
    pages = f"{start}–{end}" if end and end != start else str(start)
    return f"{location}, стр. {pages}"


def review_findings(result):
    findings = {}
    for index, item in enumerate(result["clauses"]["findings"], 1):
        findings[f"clause-{index}"] = {
            "title": item["status"].capitalize(),
            "explanation": item["explanation"],
            **{
                side: {"location": with_pdf_page(f'П. {item[side]["clause_number"]}', item[side]),
                       "quote": item[side]["quote"]}
                for side in ("template", "carrier") if side in item
            },
        }
    terms = result.get("payment") or {}
    if terms.get("status") == "отличается":
        payment_sources = {side: terms[side]["source"] for side in ("template", "carrier")}
        same_clause = any(
            all(source["clause_number"] and item.get(side, {}).get("clause_number") == source["clause_number"]
                for side, source in payment_sources.items())
            for item in result["clauses"]["findings"]
        )
        if not same_clause:
            findings["payment"] = {
                "title": "Оплата выполненного рейса",
                "explanation": "Условия оплаты отличаются.",
                **{
                    side: {
                        "location": with_pdf_page(f'П. {source["clause_number"]}' if source["clause_number"]
                                                  else f'Блок документа {source["body_position"]}', source),
                        "quote": source["quote"],
                    }
                    for side, source in payment_sources.items()
                },
            }
    for index, item in enumerate(result.get("missing_appendices", []), 1):
        findings[f"appendix-{index}"] = {
            "title": "Отсутствует приложение",
            "explanation": item["title"],
            "carrier": {"location": with_pdf_page("Ссылка в договоре", item["references"][0]),
                        "quote": "\n".join(ref["quote"] for ref in item["references"])},
        }
    for index, item in enumerate(result.get("party_details", []), 1):
        if item["status"] == "совпадает":
            continue
        findings[f"details-{index}"] = {
            "title": f'{item["role"]} · {item["field"]}',
            "explanation": item["status"],
            **{
                side: {"location": "Реквизиты", "quote": ", ".join(dict.fromkeys(row["value"] for row in item[side]))}
                for side in ("template", "carrier") if item[side]
            },
        }
    return findings


def decision_review(connection, job_id, session_id, *, raw=False):
    row = connection.execute(
        """SELECT j.state, j.result, c.contract_name FROM chat_reviews c
           JOIN jobs j ON j.id=c.review_id WHERE c.session_id=? AND j.id=?""",
        (session_id, job_id),
    ).fetchone()
    if row is None:
        raise HTTPException(404, "Проверка не найдена в этом диалоге.")
    if row["state"] == "template_mismatch":
        raise HTTPException(409, "Проверьте выбранный шаблон перед согласованием замечаний.")
    if row["state"] != "completed":
        raise HTTPException(409, "Дождитесь окончания проверки.")
    result = json.loads(row["result"])
    return row["contract_name"], result if raw else review_findings(result)


def latest_decisions(connection, job_id):
    return {
        row["finding_id"]: dict(row) for row in connection.execute(
            """SELECT id, finding_id, decision, comment FROM review_decisions WHERE id IN
               (SELECT MAX(id) FROM review_decisions WHERE review_id=? GROUP BY finding_id)
               ORDER BY id""", (job_id,)
        )
    }


def fingerprint(value):
    return hashlib.sha256(" ".join(value.split()).encode()).hexdigest()


def refused_findings(connection, job_id):
    return {row["finding_id"] for row in connection.execute(
        "SELECT finding_id FROM review_refusals WHERE review_id=?", (job_id,)
    )}


def revision_key(result, finding_id, finding):
    if finding_id.startswith("clause-"):
        source = result["clauses"]["findings"][int(finding_id.split("-")[1]) - 1]
        if "template" in source:
            return ("template-clause", source["template"]["body_position"])
        return ("added-clause", fingerprint(finding["carrier"]["quote"]))
    return (finding_id.partition("-")[0], finding["title"], finding["explanation"] if finding_id.startswith("appendix-") else "")


def compare_revision(current, parent_id, previous, decisions, refusals):
    before, after = review_findings(previous), review_findings(current)
    old_keys = {key: revision_key(previous, key, item) for key, item in before.items()}
    new_keys = {key: revision_key(current, key, item) for key, item in after.items()}
    matched = set()
    items = []
    for key, old in before.items():
        anchor = old_keys[key]
        candidates = [new_id for new_id, new_key in new_keys.items() if new_key == anchor]
        current_id = None
        status = "verify"
        if len(candidates) == 1 and list(old_keys.values()).count(anchor) == 1:
            current_id = candidates[0]
            matched.add(current_id)
            new = after[current_id]
            old_value = old.get("carrier", {}).get("quote", "")
            new_value = new.get("carrier", {}).get("quote", "")
            status = "unchanged" if fingerprint(old_value) == fingerprint(new_value) and old["title"] == new["title"] else "changed"
            if anchor[0] == "template-clause" and old_value and new_value and old["title"] == new["title"]:
                sources = [
                    report["clauses"]["findings"][int(finding_id.split("-")[1]) - 1]["carrier"]
                    for report, finding_id in ((previous, key), (current, current_id))
                ]
                hashes = [
                    report.get("clause_fingerprints", {}).get("carrier", {}).get(str(source["body_position"]))
                    for report, source in zip((previous, current), sources)
                ]
                if all(hashes):
                    status = "unchanged" if hashes[0] == hashes[1] else "changed"
        elif not candidates:
            if anchor[0] == "template-clause":
                clause = str(anchor[1])
                hashes = current["clause_fingerprints"]
                body = hashes["template"].get(clause)
                matches = [match for match in current["clauses"].get("clause_matches", [])
                           if str(match["template"]["body_position"]) == clause]
                old_source = previous["clauses"]["findings"][int(key.split("-")[1]) - 1].get("carrier")
                old_hash = previous.get("clause_fingerprints", {}).get("carrier", {}).get(
                    str(old_source["body_position"]) if old_source else ""
                )
                equivalent = len(matches) == 1 and (
                    old_source is None or (
                        fingerprint(old["carrier"]["quote"]) != fingerprint(matches[0]["carrier"]["quote"])
                        and (old_hash is None or old_hash != hashes["carrier"].get(str(matches[0]["carrier"]["body_position"])))
                    )
                )
                uncertain = any(str(item["source"]["body_position"]) == clause and item["side"] == "template"
                                for item in current["clauses"].get("manual_review", []))
                exact = (body and list(hashes["template"].values()).count(body) == 1
                         and list(hashes["carrier"].values()).count(body) == 1
                         and body != old_hash
                         and (old_source is None or fingerprint(old["carrier"]["quote"]) != body))
                if not uncertain and (equivalent or exact):
                    status = "resolved"
            elif anchor[0] == "appendix" and old["explanation"] in current["appendix_titles"]:
                status = "resolved"
            elif anchor[0] == "details" and any(
                f'{item["role"]} · {item["field"]}' == old["title"] and item["status"] == "совпадает"
                for item in current["party_details"]
            ):
                status = "resolved"
            elif anchor[0] == "payment" and current["payment"]["status"] == "совпадает":
                status = "resolved"
        items.append({
            "status": status, "current_id": current_id, "previous": old,
            "previous_decision": decisions.get(key), "previous_refusal": key in refusals,
        })
    items.extend({"status": "new", "current_id": key} for key in after if key not in matched)
    return {"previous_id": parent_id, "previous_name": previous.get("contract_name"), "items": items}


@app.post("/api/reviews/{job_id}/revisions", status_code=202)
async def submit_revision(
    job_id: str, request: Request, session_id: str = Query(min_length=1, max_length=200),
    name: str = Query(default="Новая редакция.docx", min_length=1, max_length=200),
):
    content = await uploaded(request, name)
    next_id, next_session = uuid4().hex, str(uuid4())
    with closing(db()) as connection, connection:
        connection.execute("BEGIN IMMEDIATE")
        previous_name, previous = decision_review(connection, job_id, session_id, raw=True)
        previous.setdefault("contract_name", previous_name)
        template = connection.execute("SELECT * FROM templates WHERE version=?", (previous["template_version"],)).fetchone()
        if template is None:
            raise HTTPException(409, "Версия шаблона предыдущей проверки недоступна.")
        snapshot = (job_id, previous, latest_decisions(connection, job_id), refused_findings(connection, job_id))
        connection.execute("INSERT INTO jobs(id,state) VALUES (?, 'queued')", (next_id,))
        connection.execute("INSERT INTO chat_reviews VALUES (?,?,?)", (next_session, next_id, name))
    worker.submit(process_job, next_id, dict(template), content, name, snapshot)
    return {"id": next_id, "session_id": next_session}


@app.put("/api/reviews/{job_id}/refusals/{finding_id}")
def save_carrier_refusal(job_id: str, finding_id: str, request: CarrierRefusal):
    with closing(db()) as connection, connection:
        connection.execute("BEGIN IMMEDIATE")
        _, findings = decision_review(connection, job_id, request.session_id)
        if finding_id not in findings:
            raise HTTPException(404, "Замечание не найдено.")
        if request.refused:
            connection.execute("INSERT OR IGNORE INTO review_refusals VALUES (?,?)", (job_id, finding_id))
        else:
            connection.execute("DELETE FROM review_refusals WHERE review_id=? AND finding_id=?", (job_id, finding_id))
    return {"refused": request.refused}


@app.get("/api/reviews/{job_id}/decisions")
def get_review_decisions(job_id: str, session_id: str = Query(min_length=1, max_length=200)):
    with closing(db()) as connection:
        _, result = decision_review(connection, job_id, session_id, raw=True)
        findings = review_findings(result)
        decisions = latest_decisions(connection, job_id)
        refusals = refused_findings(connection, job_id)
        draft = connection.execute(
            "SELECT decision_ids, body FROM review_drafts WHERE review_id=?", (job_id,)
        ).fetchone()
    decision_ids = [item["id"] for item in decisions.values() if item["decision"] == "propose"]
    return {
        "findings": [{"id": key, **item, "saved": decisions.get(key), "refused": key in refusals}
                     for key, item in findings.items()],
        "revision": result.get("revision"),
        "draft": draft["body"] if draft and json.loads(draft["decision_ids"]) == decision_ids else None,
    }


@app.put("/api/reviews/{job_id}/decisions/{finding_id}")
def save_review_decision(job_id: str, finding_id: str, request: FindingDecision):
    with closing(db()) as connection, connection:
        connection.execute("BEGIN IMMEDIATE")
        _, findings = decision_review(connection, job_id, request.session_id)
        if finding_id not in findings:
            raise HTTPException(404, "Замечание не найдено.")
        comment = request.comment.strip()
        previous = latest_decisions(connection, job_id).get(finding_id)
        if previous and (previous["decision"], previous["comment"]) == (request.decision, comment):
            return previous
        cursor = connection.execute(
            "INSERT INTO review_decisions(review_id,finding_id,decision,comment) VALUES (?,?,?,?)",
            (job_id, finding_id, request.decision, comment),
        )
        return {"id": cursor.lastrowid, "finding_id": finding_id, "decision": request.decision, "comment": comment}


@app.post("/api/reviews/{job_id}/draft")
def prepare_review_draft(job_id: str, request: ReviewSession):
    with closing(db()) as connection:
        name, findings = decision_review(connection, job_id, request.session_id)
        selected = [item for item in latest_decisions(connection, job_id).values() if item["decision"] == "propose"]
    if not selected:
        raise HTTPException(409, "Выберите редакцию шаблона хотя бы для одного замечания.")
    payload = {
        "action": "draft",
        "contractName": name,
        "findings": [{**findings[item["finding_id"]], "comment": item["comment"]} for item in selected],
    }
    body = call_review_workflow(payload, "Не удалось подготовить черновик. Попробуйте ещё раз.").strip()
    if not body:
        raise HTTPException(502, "Модель вернула пустой черновик.")
    ids = [item["id"] for item in selected]
    with closing(db()) as connection, connection:
        connection.execute("BEGIN IMMEDIATE")
        decision_review(connection, job_id, request.session_id)
        current = [item["id"] for item in latest_decisions(connection, job_id).values() if item["decision"] == "propose"]
        if ids != current:
            raise HTTPException(409, "Предложения изменились. Подготовьте черновик ещё раз.")
        connection.execute(
            """INSERT INTO review_drafts VALUES (?,?,?)
               ON CONFLICT(review_id) DO UPDATE SET decision_ids=excluded.decision_ids, body=excluded.body""",
            (job_id, json.dumps(ids), body),
        )
    return {"draft": body}


@app.post("/api/review-chat")
def ask_review_agent(question: ReviewQuestion):
    if not question.message.strip():
        raise HTTPException(422, "Введите вопрос по договору.")
    review = chat_review(question.session_id, section="summary")
    if review["state"] == "missing":
        raise HTTPException(404, "Сначала загрузите договор на проверку.")
    if review["state"] == "template_mismatch":
        return {"output": "Нужно проверить шаблон. " + template_fit_message(review["template_fit"]) +
                "\n\n[Посмотреть диагностику](" + review["redirect"] + ")"}
    return {"output": call_review_workflow(
        {"sessionId": question.session_id, "chatInput": question.message},
        "Чат недоступен. Попробуйте отправить вопрос ещё раз.",
    )}


def call_review_workflow(payload, error_message):
    request = UpstreamRequest(
        "http://127.0.0.1:5678/webhook/contract-review-page-chat",
        data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json", "X-Contract-Chat-Token": chat_bridge_token()},
        method="POST",
    )
    try:
        with urlopen(request, timeout=120) as response:
            result = json.load(response)
    except (URLError, TimeoutError, ValueError) as error:
        logger.warning("n8n review chat request failed: %s", error)
        raise HTTPException(502, error_message) from error
    if not isinstance(result, dict) or not isinstance(result.get("output"), str):
        raise HTTPException(502, error_message)
    return result["output"]


def page(title, body, refresh=False):
    reload = '<meta http-equiv="refresh" content="4">' if refresh else ""
    return HTMLResponse(
        f'<!doctype html><html lang="ru"><meta charset="utf-8">{reload}'
        f'<meta name="viewport" content="width=device-width, initial-scale=1">'
        f'<title>{title}</title><main style="max-width:72ch;margin:3rem auto;padding:0 1rem;'
        f'font:17px system-ui"><h1>{title}</h1>{body}</main></html>'
    )


@app.get("/template")
def template_page():
    row = active_template()
    word = f'{escape(row["name"])} · версия {row["version"][:12]}' if row else "Шаблон ещё не загружен."
    rates = rate_template_status()
    current = (f'{escape(rates["name"])} · версия {rates["version"][:12]}' if rates
               else "Согласованные тарифы ещё не загружены.")
    return page("Шаблоны проверки", f"""
<h2>Шаблон договора</h2><p>{word}</p>
<h2>Согласованные тарифы Excel</h2><p id="active-rates">{current}</p>
<p>Новые пакеты используют активную версию на момент создания.</p>
<form id="rates-template">
  <label>Таблица тарифов XLSX <input type="file" name="file" accept=".xlsx" required></label>
  <p><button type="submit">Сохранить и сделать активной</button></p>
</form>
<p id="template-status" role="status"></p>
<p><a href="http://localhost:5678/form/contract-review-templates">Загрузить шаблон договора и тарифы через n8n</a></p>
<p><a href="http://localhost:5678/form/contract-agent-batch">Открыть пакетную форму n8n</a></p>
""" + r"""
<script>
const form = document.querySelector('#rates-template');
const statusText = document.querySelector('#template-status');
form.addEventListener('submit', async event => {
  event.preventDefault();
  const file = form.elements.file.files[0];
  const button = form.querySelector('button');
  if (file.size > 2 * 1024 * 1024) {
    statusText.textContent = 'Допустимый размер XLSX — 2 МиБ.';
    return;
  }
  button.disabled = true;
  statusText.textContent = 'Сохраняем тарифы…';
  try {
    const content = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(',')[1]);
      reader.onerror = () => reject(new Error('Не удалось прочитать XLSX.'));
      reader.readAsDataURL(file);
    });
    const response = await fetch('/api/rate-template', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({name: file.name, content}),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(typeof data.detail === 'string' ? data.detail : 'Не удалось сохранить тарифы.');
    document.querySelector('#active-rates').textContent = data.name + ' · версия ' + data.version.slice(0, 12);
    statusText.textContent = 'Версия активна.';
    form.reset();
  } catch (error) {
    statusText.textContent = error.message;
  } finally {
    button.disabled = false;
  }
});
</script>
""")


@app.get("/batch", response_class=HTMLResponse)
def batch_page():
    with closing(db()) as connection:
        templates = word_templates(connection)
        rates = active_rates(connection)
    options = "".join(
        f'<option value="{escape(row["version"])}">{escape(row["name"])} · {row["version"][:12]}</option>'
        for row in templates
    )
    rate_hint = (f'Активная таблица: {escape(rates["name"])} · версия {rates["version"][:12]}' if rates
                 else 'Сохраните согласованные тарифы или приложите обе таблицы.')
    return page("Пакетная проверка договоров", f"""
<form id="upload">
  <label>Общий шаблон <select name="template" required>{options}</select></label>
  <label>Договоры DOCX или PDF <input type="file" name="files" accept=".docx,.pdf" multiple required></label>
  <details>
    <summary>Сверить тарифы из Excel</summary>
    <p>{rate_hint} <a href="/template">Шаблоны проверки</a></p>
    <label>Согласованные тарифы для разовой сверки <input type="file" name="approved" accept=".xlsx"></label>
    <label>Тарифы перевозчика <input type="file" name="vendor" accept=".xlsx"></label>
  </details>
  <button type="submit">Создать пакет и проверить</button>
</form>
""" + r"""
<style>
  form label { display: block; margin: .75rem 0; }
  form input, form select, form textarea, button { font: inherit; }
  button { padding: .45rem .7rem; }
  #documents details { border: 1px solid #cbd5e1; border-radius: .45rem; margin: .5rem 0; padding: .75rem; }
  #documents summary { cursor: pointer; overflow-wrap: anywhere; }
  #rates { border: 1px solid #cbd5e1; border-radius: .45rem; margin: 1rem 0; padding: .85rem; }
  #rates summary { cursor: pointer; font-weight: 600; }
  #ask { margin-top: 1.5rem; }
  #ask textarea { display: block; width: 100%; box-sizing: border-box; min-height: 8rem; margin-top: .45rem;
    padding: .65rem; resize: vertical; line-height: 1.5; border: 1px solid #94a3b8; border-radius: .35rem; }
  #answer { white-space: pre-wrap; overflow-wrap: anywhere; }
</style>
<p id="status" role="status"></p>
<section id="results" hidden>
  <a href="/batch">Создать новый пакет</a>
  <p id="overview"></p>
  <details id="rates" open>
    <summary id="rates-heading">Сверка тарифов</summary>
    <div id="rate-findings"></div>
    <form id="attach-rates" hidden>
      <p id="rate-template-basis"></p>
      <label>Согласованные тарифы для разовой сверки <input type="file" name="approved" accept=".xlsx" required></label>
      <label>Тарифы перевозчика XLSX <input type="file" name="vendor" accept=".xlsx" required></label>
      <button type="submit">Сверить Excel для этого пакета</button>
    </form>
    <p id="rates-status" role="status" hidden></p>
  </details>
  <div id="documents"></div>
  <form id="ask">
    <label>Вопрос по пакету
      <textarea name="question" rows="5" placeholder="Где изменился срок оплаты? Укажи файл и пункт договора." required maxlength="2000"></textarea>
    </label>
    <button>Спросить агента</button>
  </form>
  <p id="answer" role="status"></p>
</section>
<script>
const upload = document.querySelector('#upload');
const attachRates = document.querySelector('#attach-rates');
const ratesStatus = document.querySelector('#rates-status');
const ask = document.querySelector('#ask');
const statusText = document.querySelector('#status');
const documents = document.querySelector('#documents');
const answer = document.querySelector('#answer');
const states = {waiting: 'Ожидает загрузки', queued: 'В очереди', running: 'Проверяется',
                completed: 'Готово', failed: 'Ошибка', template_mismatch: 'Нужно проверить шаблон'};
const requestedBatch = new URLSearchParams(location.search).get('id');
let batchId = /^[a-f0-9]{32}$/.test(requestedBatch ?? '') ? requestedBatch : null;
let poll;
let previous = '';
let refreshVersion = 0;

async function jsonRequest(path, options) {
  const response = await fetch(path, options);
  const data = await response.json();
  if (!response.ok) throw new Error(typeof data.detail === 'string' ? data.detail : 'Запрос не выполнен.');
  return data;
}

function element(tag, text) {
  const node = document.createElement(tag);
  node.textContent = text;
  return node;
}

function renderRates(rates, reference) {
  const panel = document.querySelector('#rates');
  const body = panel.querySelector('#rate-findings');
  const finished = rates.state === 'completed';
  const labelsByState = {not_uploaded: 'Excel не загружен', incomplete: 'Нужны обе таблицы',
    invalid: 'Ошибка в таблице', completed: 'Готово'};
  const suffix = finished ? ' · расхождений ' + rates.findings.length : '';
  panel.querySelector('#rates-heading').textContent = 'Сверка тарифов · ' + labelsByState[rates.state] + suffix;
  attachRates.hidden = finished;
  attachRates.elements.approved.required = !reference;
  document.querySelector('#rate-template-basis').textContent = reference
    ? 'Согласованные тарифы пакета: ' + reference.name + ' · версия ' + reference.version.slice(0, 12)
    : 'Для этого пакета нужно приложить обе таблицы.';
  if (location.hash === '#rates') panel.open = true;
  if (rates.state !== 'completed') {
    body.replaceChildren(element('p', rates.message ??
      (reference ? 'Добавьте тарифы перевозчика, чтобы запустить сверку.' : 'Добавьте обе таблицы, чтобы запустить сверку.')));
    return;
  }
  const labels = {changed: 'Тариф изменён', missing: 'Нет у перевозчика',
    added: 'Нет в согласованной таблице', incompatible: 'Различаются валюта, единица расчёта или НДС'};
  const list = document.createElement('ul');
  for (const finding of rates.findings) {
    const route = finding.vendor ?? finding.approved;
    const item = element('li', route.origin + ' → ' + route.destination + ' · ' + route.vehicle);
    item.append(element('p', labels[finding.status] +
      (finding.difference === null ? '' : ' · разница ' + finding.difference + ' ' + route.currency)));
    for (const [label, side] of [['Согласовано', finding.approved], ['Перевозчик', finding.vendor]]) {
      if (!side) continue;
      item.append(element('p', label + ' · ' + side.rate + ' ' + side.currency + ' / ' + side.unit +
        ' · ' + side.vat + ' · ' + side.file + ', ' + side.sheet + '!' + side.cell));
    }
    list.append(item);
  }
  const version = rates.approved_version ? ' · версия ' + rates.approved_version.slice(0, 12) : '';
  body.replaceChildren(element('p', 'Согласовано: ' + rates.approved_file + version),
    element('p', 'Перевозчик: ' + rates.vendor_file),
    element('p', 'Совпало ' + rates.matched + ' · расхождений ' + rates.findings.length), list);
}

function encodeRates(file) {
  if (!file) return null;
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({name: file.name, content: reader.result.split(',')[1]});
    reader.onerror = () => reject(new Error('Не удалось прочитать ' + file.name));
    reader.readAsDataURL(file);
  });
}

function ratesTooLarge(approved, vendor) {
  return [approved, vendor].some(file => file && file.size > 2 * 1024 * 1024);
}

async function saveRates(id, approved, vendor) {
  return jsonRequest('/api/batches/' + id + '/rates', {
    method: 'PUT', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({approved: await encodeRates(approved), vendor: await encodeRates(vendor)}),
  });
}

async function refresh() {
  const current = batchId;
  const version = ++refreshVersion;
  clearTimeout(poll);
  try {
    const data = await jsonRequest('/api/batches/' + current);
    if (batchId !== current || version !== refreshVersion) return;
    const done = data.items.filter(item => item.state === 'completed').length;
    const failed = data.items.filter(item => item.state === 'failed').length;
    const mismatches = data.items.filter(item => item.state === 'template_mismatch').length;
    const count = data.items.reduce((sum, item) => sum + item.findings.length, 0);
    const signature = JSON.stringify(data);
    if (signature !== previous) {
      renderRates(data.rates, data.rate_template);
      const open = documents.querySelector('details[open]')?.dataset.index;
      documents.replaceChildren();
      data.items.forEach((item, index) => {
        const card = document.createElement('details');
        card.dataset.index = String(index);
        card.open = open === card.dataset.index;
        card.setAttribute('name', 'contracts');
        card.append(element('summary', item.name + ' · ' + states[item.state] +
                    (item.state === 'completed' ? ' · замечаний ' + item.findings.length : '')));
        if (item.state === 'completed') {
          card.append(element('p', 'Оплата — ' + (item.payment ?? 'нет результата') + ' · ручная сверка пунктов — ' + item.manual));
          const list = document.createElement('ul');
          for (const finding of item.findings) {
            list.append(element('li', finding.title + ' · ' + (finding.carrier ?? finding.template).location + ' · ' + finding.explanation));
          }
          if (list.children.length) card.append(list);
          const report = element('a', 'Открыть отчёт');
          report.href = '/reviews/' + item.review_id;
          const dialog = element('a', 'Решения по замечаниям');
          dialog.href = '/review#' + new URLSearchParams({review: item.review_id, session: item.session_id});
          card.append(report, ' · ', dialog);
        } else if (item.state === 'template_mismatch') {
          const fit = item.template_fit;
          card.append(element('p', 'Найдено близких пар: ' + fit.matched_pairs + ' из ' +
            fit.template_count + ' пунктов шаблона и ' + fit.matched_pairs + ' из ' +
            fit.carrier_count + ' пунктов договора.'));
          const diagnosis = element('a', 'Посмотреть диагностику');
          diagnosis.href = '/reviews/' + item.review_id;
          card.append(diagnosis);
        } else if (item.state === 'failed') {
          card.append(element('p', item.error ?? 'Проверка завершилась ошибкой.'));
        }
        documents.append(card);
      });
      previous = signature;
    }
    document.querySelector('#overview').textContent = 'Шаблон ' + data.template.name +
      ' · версия ' + data.template.version.slice(0, 12) +
      ' · готово ' + done + ' из ' + data.items.length + ' · ошибок ' + failed + ' · замечаний ' + count +
      (mismatches ? ' · проверить шаблон ' + mismatches : '');
    document.querySelector('#results').hidden = false;
    upload.hidden = true;
    ask.querySelector('button').disabled = !done && !mismatches && data.rates.state === 'not_uploaded';
    if (done + failed + mismatches < data.items.length) poll = setTimeout(refresh, 3000);
  } catch (error) {
    if (batchId !== current || version !== refreshVersion) return;
    if (document.querySelector('#results').hidden) upload.hidden = false;
    statusText.textContent = error.message;
  }
}

upload.addEventListener('submit', async event => {
  event.preventDefault();
  const files = Array.from(upload.elements.files.files);
  if (!files.length || files.length > 10) {
    statusText.textContent = 'Выберите от 1 до 10 файлов.';
    return;
  }
  const button = upload.querySelector('button');
  const approved = upload.elements.approved.files[0];
  const vendor = upload.elements.vendor.files[0];
  if (ratesTooLarge(approved, vendor)) {
    statusText.textContent = 'Допустимый размер каждого XLSX — 2 МиБ.';
    return;
  }
  button.disabled = true;
  statusText.textContent = 'Загружаем файлы…';
  try {
    const batch = await jsonRequest('/api/batches', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({files: files.map(file => file.name), template_version: upload.elements.template.value}),
    });
    clearTimeout(poll);
    batchId = batch.id;
    previous = '';
    answer.textContent = '';
    history.replaceState(null, '', '?id=' + batchId);
    if (approved || vendor) {
      try {
        await saveRates(batchId, approved, vendor);
      } catch (error) {
        statusText.textContent = error.message;
      }
    }
    for (const [index, file] of files.entries()) {
      try {
        await jsonRequest('/api/batches/' + batchId + '/files/' + index, {method: 'PUT', body: file});
      } catch (error) {
        statusText.textContent = file.name + ' · ' + error.message;
      }
    }
    if (statusText.textContent === 'Загружаем файлы…') statusText.textContent = '';
    void refresh();
  } catch (error) {
    statusText.textContent = error.message;
  } finally {
    button.disabled = false;
  }
});

attachRates.addEventListener('submit', async event => {
  event.preventDefault();
  const current = batchId;
  const approved = attachRates.elements.approved.files[0];
  const vendor = attachRates.elements.vendor.files[0];
  const button = attachRates.querySelector('button');
  ratesStatus.hidden = false;
  if (ratesTooLarge(approved, vendor)) {
    ratesStatus.textContent = 'Допустимый размер каждого XLSX — 2 МиБ.';
    return;
  }
  button.disabled = true;
  ratesStatus.textContent = 'Сверяем тарифы…';
  try {
    await saveRates(current, approved, vendor);
    ratesStatus.hidden = true;
    await refresh();
  } catch (error) {
    ratesStatus.textContent = error.message;
  } finally {
    button.disabled = false;
  }
});

ask.addEventListener('submit', async event => {
  event.preventDefault();
  const current = batchId;
  const button = ask.querySelector('button');
  button.disabled = true;
  answer.textContent = 'Агент проверяет результаты…';
  try {
    const storageKey = 'contractBatchChat.' + current;
    let sessionId = sessionStorage.getItem(storageKey);
    if (!sessionId) {
      sessionId = crypto.randomUUID();
      sessionStorage.setItem(storageKey, sessionId);
    }
    const result = await jsonRequest('/api/batches/' + current + '/chat', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({message: ask.elements.question.value, session_id: sessionId}),
    });
    if (current !== batchId) return;
    answer.replaceChildren();
    let offset = 0;
    for (const match of result.output.matchAll(/\[([^\]\n]+)\]\(http:\/\/localhost:8765(\/reviews\/[a-f0-9]{32})\)/g)) {
      const link = element('a', match[1]);
      link.href = match[2];
      answer.append(result.output.slice(offset, match.index), link);
      offset = match.index + match[0].length;
    }
    answer.append(result.output.slice(offset));
  } catch (error) {
    if (current === batchId) answer.textContent = error.message;
  } finally {
    button.disabled = false;
  }
});
if (batchId) {
  upload.hidden = true;
  void refresh();
}
</script>
""")


@app.get("/review", response_class=HTMLResponse)
def review_upload_page():
    template = active_template()
    if template is None:
        return page("Проверка договора", '<p>Сначала загрузите шаблон заказчика.</p>')
    return page("Проверка договора", f'<p>Активный шаблон · {escape(template["name"])}</p>' + r"""
<style>
  #review-progress, #review-chat { border: 1px solid #cbd5e1; border-radius: .75rem; padding: 1.25rem; }
  #review-chat button, #report-link {
    background: #1d4ed8; color: white; border: 0; border-radius: .4rem;
    padding: .65rem 1rem; font: inherit; font-weight: 650; text-decoration: none;
  }
  #review-chat button:disabled { opacity: .65; }
  #review-progress, #review-chat { margin-top: 1rem; }
  #review-status { margin: 0; }
  #report-link:not([hidden]) { display: inline-block; margin-top: .75rem; }
  #review-chat h2 { margin-top: 0; font-size: 1.2rem; }
  #review-chat input { box-sizing: border-box; width: 100%; padding: .55rem; margin: .4rem 0 .65rem; font: inherit; }
  #chat-log p { white-space: pre-wrap; overflow-wrap: anywhere; }
  #review-decisions { margin-top: 1.5rem; border-top: 1px solid #cbd5e1; padding-top: 1rem; }
  #review-decisions article { border: 1px solid #cbd5e1; border-radius: .5rem; padding: 1rem; margin-bottom: .75rem; }
  #review-decisions h3 { margin: 0 0 .5rem; font-size: 1rem; }
  #review-decisions blockquote { margin: .65rem 0; padding-left: .75rem; border-left: 2px solid #cbd5e1; white-space: pre-wrap; }
  #review-decisions textarea { display: block; box-sizing: border-box; width: 100%; padding: .5rem; font: inherit; margin: .35rem 0 .75rem; }
  #review-decisions .choices { display: flex; flex-wrap: wrap; gap: .4rem; }
  #review-decisions .choices button[aria-pressed="false"] { background: #e8eef8; color: #17315b; }
  #draft-box { margin-top: 1rem; }
  #revision-form { margin-top: 1rem; }
  #revision-form input { display: block; margin: .5rem 0; }
  #revision-summary li { margin: .6rem 0; }
  #review-decisions .refusal input { width: auto; margin-right: .5rem; }
</style>
<p><a href="http://localhost:5678/form/contract-agent-review">Проверить другой договор</a> · <a href="/batch">Проверить пакет договоров</a></p>
<div id="review-progress" hidden>
  <p id="review-status" role="status"></p>
  <a id="report-link" target="_blank" rel="opener" hidden>Открыть отчёт в новой вкладке</a>
  <p id="previous-review" hidden><a>Предыдущая редакция</a></p>
  <form id="revision-form" hidden>
    <label for="revision-file">Новая редакция договора</label>
    <input id="revision-file" type="file" accept=".docx,.pdf" required>
    <button type="submit">Загрузить новую редакцию</button>
    <p id="revision-status" role="status"></p>
  </form>
</div>
<section id="review-chat" hidden>
  <h2>Чат по договору</h2>
  <div id="chat-log" role="log" aria-label="Переписка по договору"></div>
  <form id="chat-form">
    <label for="chat-question">Вопрос по договору</label>
    <input id="chat-question" type="text" maxlength="2000" required placeholder="Что изменилось в оплате?">
    <button type="submit">Отправить</button>
  </form>
  <p id="chat-status" role="status"></p>
  <section id="review-decisions" hidden>
    <h2>Решения по замечаниям</h2>
    <section id="revision-summary" hidden>
      <h3>Что изменилось с прошлой редакции</h3>
      <ul></ul>
    </section>
    <div id="finding-cards"></div>
    <button id="make-draft" type="button" disabled>Подготовить черновик</button>
    <p id="decision-status" role="status"></p>
    <div id="draft-box" hidden>
      <label for="draft-text">Черновик письма перевозчику</label>
      <textarea id="draft-text" rows="12"></textarea>
      <button id="copy-draft" type="button">Скопировать письмо</button>
    </div>
  </section>
</section>
<script>
const progress = document.querySelector('#review-progress');
const status = document.querySelector('#review-status');
const reportLink = document.querySelector('#report-link');
const chat = document.querySelector('#review-chat');
const chatLog = document.querySelector('#chat-log');
const chatForm = document.querySelector('#chat-form');
const question = document.querySelector('#chat-question');
const chatSubmit = chatForm.querySelector('button');
const chatStatus = document.querySelector('#chat-status');
const reviewStorageKey = 'contract-review.current';
const revisionForm = document.querySelector('#revision-form');
const revisionFile = document.querySelector('#revision-file');
const revisionStatus = document.querySelector('#revision-status');
const previousReview = document.querySelector('#previous-review');
let sessionId;
let reviewId;
const decisionsPanel = document.querySelector('#review-decisions');
const cards = document.querySelector('#finding-cards');
const decisionStatus = document.querySelector('#decision-status');
const makeDraft = document.querySelector('#make-draft');
const draftBox = document.querySelector('#draft-box');
const draftText = document.querySelector('#draft-text');
const decisions = new Map();
const decisionLabels = {accept: 'Принято', propose: 'Предложена правка', discuss: 'На обсуждении'};
const revisionLabels = {
  unchanged: 'Замечание осталось', changed: 'Формулировка изменилась',
  resolved: 'Правка внесена', verify: 'Нужно сверить вручную', new: 'Новое замечание',
};
let pendingChanges = 0;
let drafting = false;

async function reviewAction(path, method, fields = {}) {
  const url = '/api/reviews/' + encodeURIComponent(reviewId) + '/' + path;
  const options = method === 'GET' ? {cache: 'no-store'} : {
    method, headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({session_id: sessionId, ...fields}),
  };
  const response = await fetch(method === 'GET' ? url + '?' + new URLSearchParams({session_id: sessionId}) : url, options);
  const result = await response.json();
  if (!response.ok) throw new Error(typeof result.detail === 'string' ? result.detail : 'Не удалось выполнить действие.');
  return result;
}

function showDraft(value) {
  draftText.value = value ?? '';
  draftBox.hidden = !value;
}

function updateDraftButton() {
  makeDraft.disabled = drafting || pendingChanges > 0 || !Array.from(decisions.values()).includes('propose');
}

function makeFindingCard(finding) {
  const card = document.createElement('article');
  card.id = finding.id;
  const title = document.createElement('h3');
  title.textContent = finding.title + ' · ' + (finding.carrier ?? finding.template).location;
  const explanation = document.createElement('p');
  explanation.textContent = finding.explanation;
  card.append(title, explanation);
  for (const [side, label] of [['template', 'Шаблон'], ['carrier', 'Договор']]) {
    const evidence = finding[side];
    if (!evidence) continue;
    const quote = document.createElement('blockquote');
    const source = document.createElement('strong');
    source.textContent = label + ' · ' + evidence.location;
    quote.append(source, '\n' + evidence.quote);
    card.append(quote);
  }
  const label = document.createElement('label');
  const comment = document.createElement('textarea');
  comment.id = 'comment-' + finding.id;
  comment.rows = 2;
  comment.maxLength = 2000;
  comment.value = finding.saved?.comment ?? '';
  label.htmlFor = comment.id;
  label.textContent = 'Комментарий';
  const choices = document.createElement('div');
  choices.className = 'choices';
  const saved = document.createElement('p');
  saved.setAttribute('role', 'status');
  if (finding.saved) decisions.set(finding.id, finding.saved.decision);
  const labels = {
    accept: 'Принять',
    propose: finding.template ? 'Предложить редакцию шаблона' : 'Предложить правку',
    discuss: 'На обсуждение',
  };
  for (const [decision, label] of Object.entries(labels)) {
    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.setAttribute('aria-pressed', String(decisions.get(finding.id) === decision));
    button.addEventListener('click', async () => {
      const buttons = Array.from(choices.children);
      buttons.forEach(choice => choice.disabled = true);
      comment.disabled = true;
      pendingChanges++;
      updateDraftButton();
      try {
        const item = await reviewAction('decisions/' + finding.id, 'PUT', {decision, comment: comment.value});
        decisions.set(finding.id, item.decision);
        buttons.forEach(choice => choice.setAttribute('aria-pressed', String(choice === button)));
        if (finding.saved?.id !== item.id && (finding.saved?.decision === 'propose' || item.decision === 'propose')) {
          showDraft(null);
        }
        finding.saved = item;
        saved.textContent = 'Сохранено · ' + label;
      } catch (error) {
        saved.textContent = error.message;
      } finally {
        pendingChanges--;
        buttons.forEach(choice => choice.disabled = false);
        comment.disabled = false;
        updateDraftButton();
      }
    });
    choices.append(button);
  }
  const refusalLabel = document.createElement('label');
  refusalLabel.className = 'refusal';
  const refusal = document.createElement('input');
  refusal.type = 'checkbox';
  refusal.checked = finding.refused;
  refusalLabel.append(refusal, 'Перевозчик отказал');
  refusal.addEventListener('change', async () => {
    refusal.disabled = true;
    try {
      const result = await reviewAction('refusals/' + finding.id, 'PUT', {refused: refusal.checked});
      refusal.checked = result.refused;
      saved.textContent = 'Ответ перевозчика сохранён.';
    } catch (error) {
      refusal.checked = !refusal.checked;
      saved.textContent = error.message;
    } finally {
      refusal.disabled = false;
    }
  });
  card.append(label, comment, choices, refusalLabel, saved);
  return card;
}

function showPreviousReview(id, name) {
  previousReview.hidden = false;
  const link = previousReview.querySelector('a');
  link.href = '/review?review=' + encodeURIComponent(id);
  link.textContent = name ? 'Предыдущая редакция · ' + name : 'Предыдущая редакция';
}

function renderRevision(revision, findings) {
  showPreviousReview(revision.previous_id, revision.previous_name);
  const summary = document.querySelector('#revision-summary');
  const list = summary.querySelector('ul');
  for (const item of revision.items) {
    const current = findings.find(finding => finding.id === item.current_id);
    const evidence = current ?? item.previous;
    const row = document.createElement('li');
    const title = revisionLabels[item.status] + ' · ' + evidence.title + ' · ' +
      (evidence.carrier ?? evidence.template).location;
    if (current) {
      const link = document.createElement('a');
      link.href = '#' + current.id;
      link.textContent = title;
      row.append(link);
    } else {
      row.textContent = title;
    }
    const decision = item.previous_decision;
    if (decision || item.previous_refusal) {
      const previous = document.createElement('div');
      previous.textContent = [
        decision && 'Ранее · ' + decisionLabels[decision.decision],
        decision?.comment, item.previous_refusal && 'Перевозчик отказал',
      ].filter(Boolean).join(' · ');
      row.append(previous);
    }
    if (item.previous?.carrier && ['changed', 'verify'].includes(item.status)) {
      const details = document.createElement('details');
      const label = document.createElement('summary');
      label.textContent = 'Формулировка в предыдущей редакции';
      const quote = document.createElement('blockquote');
      quote.textContent = item.previous.carrier.location + '\n' + item.previous.carrier.quote;
      details.append(label, quote);
      row.append(details);
    }
    list.append(row);
  }
  summary.hidden = false;
}

async function loadDecisions() {
  const result = await reviewAction('decisions', 'GET');
  decisions.clear();
  cards.replaceChildren(...result.findings.map(makeFindingCard));
  if (result.revision) renderRevision(result.revision, result.findings);
  if (!result.findings.length) cards.textContent = 'Нет замечаний для согласования.';
  decisionsPanel.hidden = false;
  makeDraft.hidden = !result.findings.length;
  showDraft(result.draft);
  updateDraftButton();
  const target = location.hash.slice(1);
  if (target) document.getElementById(target)?.scrollIntoView();
}

revisionForm.addEventListener('submit', async event => {
  event.preventDefault();
  const file = revisionFile.files[0];
  if (!file) return;
  const button = revisionForm.querySelector('button');
  button.disabled = revisionFile.disabled = true;
  revisionStatus.textContent = 'Загружаем новую редакцию…';
  try {
    const params = new URLSearchParams({session_id: sessionId, name: file.name});
    const response = await fetch('/api/reviews/' + encodeURIComponent(reviewId) + '/revisions?' + params,
                                 {method: 'POST', body: file});
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || 'Не удалось загрузить редакцию.');
    const next = {sessionId: result.session_id, reviewId: result.id, previousId: reviewId};
    sessionStorage.setItem(reviewStorageKey + '.' + next.reviewId, JSON.stringify(next));
    sessionStorage.setItem(reviewStorageKey, JSON.stringify(next));
    location.assign('/review?review=' + encodeURIComponent(next.reviewId));
  } catch (error) {
    revisionStatus.textContent = error.message;
    button.disabled = revisionFile.disabled = false;
  }
});

makeDraft.addEventListener('click', async () => {
  drafting = true;
  updateDraftButton();
  decisionStatus.textContent = 'Готовим черновик…';
  try {
    const result = await reviewAction('draft', 'POST');
    showDraft(result.draft);
    decisionStatus.textContent = 'Черновик готов. Текст можно отредактировать перед копированием.';
  } catch (error) {
    decisionStatus.textContent = error.message;
  } finally {
    drafting = false;
    updateDraftButton();
  }
});

document.querySelector('#copy-draft').addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(draftText.value);
    decisionStatus.textContent = 'Письмо скопировано.';
  } catch {
    draftText.select();
    decisionStatus.textContent = 'Текст выделен. Нажмите Ctrl+C, чтобы скопировать.';
  }
});

function addChatMessage(author, message) {
  const entry = document.createElement('p');
  const label = document.createElement('strong');
  label.textContent = author;
  entry.append(label, '\n');
  let cursor = 0;
  if (author === 'Агент') {
    const links = /\[([^\]\r\n]+)\]\(http:\/\/localhost:8765(\/reviews\/[a-f0-9]{32})\)/g;
    for (const match of message.matchAll(links)) {
      entry.append(message.slice(cursor, match.index));
      const link = document.createElement('a');
      link.textContent = match[1];
      link.href = match[2];
      link.target = '_blank';
      link.rel = 'opener';
      entry.append(link);
      cursor = match.index + match[0].length;
    }
  }
  entry.append(message.slice(cursor));
  chatLog.append(entry);
}

chatForm.addEventListener('submit', async event => {
  event.preventDefault();
  const message = question.value.trim();
  if (!message || !sessionId || chatSubmit.disabled) return;
  const currentSession = sessionId;
  question.disabled = chatSubmit.disabled = true;
  chatStatus.textContent = 'Агент готовит ответ…';
  try {
    const response = await fetch('/api/review-chat', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({session_id: currentSession, message}),
    });
    const result = await response.json();
    if (currentSession !== sessionId) return;
    if (!response.ok) throw new Error(result.detail || 'Не удалось получить ответ.');
    addChatMessage('Вы', message);
    addChatMessage('Агент', result.output);
    question.value = '';
    chatStatus.textContent = '';
  } catch (error) {
    if (currentSession === sessionId) chatStatus.textContent = error.message;
  } finally {
    question.disabled = chatSubmit.disabled = false;
  }
});

async function waitForReport(id) {
  const statusUrl = '/api/reviews/' + encodeURIComponent(id);
  while (true) {
    let response;
    let review;
    try {
      response = await fetch(statusUrl, {cache: 'no-store', signal: AbortSignal.timeout(10000)});
      if (response.ok) review = await response.json();
    } catch {
      // A temporary network error can clear before the next status check.
    }
    if (response?.status === 404) throw new Error('Проверка не найдена.');
    if (review?.state === 'completed' || review?.state === 'template_mismatch') {
      const completed = review.state === 'completed';
      status.textContent = completed ? 'Отчёт готов' : 'Нужно проверить шаблон';
      reportLink.textContent = completed ? 'Открыть отчёт в новой вкладке' : 'Посмотреть диагностику';
      reportLink.href = '/reviews/' + encodeURIComponent(id);
      reportLink.hidden = false;
      return completed;
    }
    if (review?.state === 'failed') throw new Error('Проверка завершилась ошибкой. Попробуйте ещё раз.');
    status.textContent = review ? 'Договор проверяется…' : 'Связь с сервисом прервалась. Проверяем статус…';
    await new Promise(resolve => setTimeout(resolve, 3000));
  }
}

async function restoreReview() {
  const launch = new URLSearchParams(location.hash.slice(1));
  const launchSession = launch.get('session');
  const launchReview = launch.get('review');
  if (/^[a-f0-9-]{36}$/i.test(launchSession ?? '') && /^[a-f0-9]{32}$/i.test(launchReview ?? '')) {
    const value = JSON.stringify({sessionId: launchSession, reviewId: launchReview});
    sessionStorage.setItem(reviewStorageKey, value);
    sessionStorage.setItem(reviewStorageKey + '.' + launchReview, value);
    history.replaceState(null, '', location.pathname + location.search);
  }
  const requestedReview = new URLSearchParams(location.search).get('review');
  const saved = JSON.parse((requestedReview && sessionStorage.getItem(reviewStorageKey + '.' + requestedReview))
                           || sessionStorage.getItem(reviewStorageKey) || 'null');
  if (!saved?.sessionId || !saved.reviewId || (requestedReview && requestedReview !== saved.reviewId)) {
    if (requestedReview) {
      progress.hidden = false;
      status.textContent = 'Откройте диалог из вкладки, в которой запустили проверку.';
    }
    return;
  }
  sessionId = saved.sessionId;
  reviewId = saved.reviewId;
  sessionStorage.setItem(reviewStorageKey + '.' + reviewId, JSON.stringify(saved));
  if (saved.previousId) showPreviousReview(saved.previousId);
  chat.hidden = progress.hidden = false;
  status.textContent = 'Проверяем статус…';
  try {
    if (!await waitForReport(saved.reviewId)) return;
  } catch (error) {
    status.textContent = error.message;
    return;
  }
  revisionForm.hidden = false;
  try {
    await loadDecisions();
  } catch (error) {
    decisionsPanel.hidden = false;
    makeDraft.hidden = true;
    decisionStatus.textContent = error.message;
  }
}

void restoreReview();
</script>
""")


@app.get("/reviews/{job_id}")
def review_page(job_id: str):
    job = review_status(job_id)
    if job["state"] in ("queued", "running"):
        return page("Договор проверяется", "<p>Эта страница обновится автоматически.</p>", refresh=True)
    if job["state"] == "failed":
        return page("Проверка прервана", "<p>Проверьте, что локальная модель запущена, и загрузите договор снова.</p>")
    result = job["result"]
    if job["state"] == "template_mismatch":
        fit = result["template_fit"]
        coverage = f'{fit["minimum_coverage"]:.0%}'
        return page("Нужно проверить шаблон",
                    f'<p><strong>Договор</strong> {escape(result["contract_name"])}<br>'
                    f'<strong>Шаблон</strong> {escape(result["template_name"])} · версия {result["template_version"][:12]}</p>'
                    f'<p>{escape(template_fit_message(fit))}</p>'
                    f'<p>Для запуска детальной сверки требуется покрытие от {coverage} пунктов с каждой стороны.</p>'
                    '<p><a href="http://localhost:5678/form/contract-agent-batch">Выбрать шаблон для новой проверки</a></p>')
    dialog_url = f'/review?review={escape(job_id, quote=True)}'
    findings = result["clauses"]["findings"]
    items = []
    for index, finding in enumerate(findings, 1):
        quotes = "".join(
            f'<blockquote><strong>{label}, {escape(with_pdf_page("п. " + str(finding[side]["clause_number"]), finding[side]))}</strong><br>'
            f'{escape(finding[side]["quote"])}</blockquote>'
            for side, label in (("template", "Шаблон"), ("carrier", "Договор")) if side in finding
        )
        items.append(
            f'<li><strong>{escape(finding["status"])}</strong> — {escape(finding["explanation"])}{quotes}'
            f'<a href="{dialog_url}#clause-{index}">Открыть замечание в диалоге</a></li>'
        )
    details = escape(json.dumps({"payment": result["payment"], "party_details": result["party_details"],
                                 "manual_review": result["clauses"]["manual_review"]}, ensure_ascii=False, indent=2))
    appendices = []
    for index, finding in enumerate(result.get("missing_appendices", []), 1):
        references = []
        for ref in finding["references"]:
            if ref["clause_number"]:
                location = f'П. {ref["clause_number"]}'
            elif "table_number" in ref:
                location = f'Таблица {ref["table_number"]}, строка {ref["row_number"]}'
            else:
                location = f'Блок документа {ref["body_position"]}'
            references.append(f'<blockquote><strong>{escape(with_pdf_page(location, ref))}</strong><br>{escape(ref["quote"])}</blockquote>')
        appendices.append(
            f'<li><strong>{escape(finding["title"])}</strong>{"".join(references)}'
            f'<a href="{dialog_url}#appendix-{index}">Открыть замечание в диалоге</a></li>'
        )
    appendix_section = (
        f'<h2>Отсутствующие приложения</h2><ul>{"".join(appendices)}</ul>' if appendices else ""
    )
    if sources := result.get("appendix_sources"):
        appendix_items = "".join(f'<li>{escape(with_pdf_page(item["title"], item))}</li>' for item in sources)
        appendix_section += f'<h2>Приложения в PDF</h2><ul>{appendix_items}</ul>'
    if tables := result.get("pdf_tables"):
        table_items = "".join(f'<li>Таблица {item["number"]}, стр. {item["page"]} · строк {item["rows"]}</li>'
                              for item in tables)
        appendix_section += f'<details><summary>Распознанные таблицы PDF</summary><ul>{table_items}</ul></details>'
    party_items = []
    for index, item in enumerate(result["party_details"], 1):
        if item["status"] == "совпадает":
            continue
        template, contract = (
            ", ".join(dict.fromkeys(entry["value"] for entry in item[side])) or "не указано"
            for side in ("template", "carrier")
        )
        party_items.append(
            f'<li><strong>{escape(item["role"])} · {escape(item["field"])}</strong>'
            f'<br>Шаблон — {escape(template)}<br>Договор — {escape(contract)}'
            f'<br>{escape(item["status"])}<br>'
            f'<a href="{dialog_url}#details-{index}">Открыть замечание в диалоге</a></li>'
        )
    party_section = f'<h2>Замечания по реквизитам</h2><ul>{"".join(party_items)}</ul>' if party_items else ""
    matches = result["clauses"].get("clause_matches", [])
    match_items = "".join(
        f'<li>П. {escape(match["template"]["clause_number"])} шаблона → '
        f'п. {escape(match["carrier"]["clause_number"])} договора, {escape(match["kind"])}</li>'
        for match in matches
    )
    match_section = f'<h2>Совпадающие условия</h2><ul>{match_items}</ul>' if match_items else ""
    clause_report = result["clauses"]
    counts = clause_report["clause_counts"]
    manual_count = len(clause_report["manual_review"])
    structure = result.get("document_structure", {}).get("carrier")
    coverage = ""
    if structure is not None:
        coverage = (
            f'<p>В договоре распознано таблиц: {structure["tables"]}, '
            f'приложений: {structure["appendices"]}, абзацев без номера: {structure["unnumbered_paragraphs"]}. '
            'В таблицах и абзацах без номера проверяются целевые данные: оплата, реквизиты и ссылки на приложения.</p>'
        )
    contract_name = escape(result["contract_name"] or "имя файла не сохранилось")
    body = (f'<p><strong>Проверенный договор</strong> {contract_name}<br>'
            f'<strong>Шаблон</strong> {escape(result["template_name"])} · версия {result["template_version"][:12]}</p>'
            f'<h2>Сводка проверки</h2><p>Сопоставление нумерованных пунктов: '
            f'{counts["template"]} в шаблоне, {counts["carrier"]} в договоре.<br>'
            f'Замечаний по пунктам: {len(findings)}. Пунктов для ручной проверки: {manual_count}.<br>'
            f'Замечаний по реквизитам: {len(party_items)}. '
            f'Отсутствующих приложений: {len(appendices)}.</p>{coverage}'
            f'<p><a href="{dialog_url}#review-decisions">Перейти к решениям по замечаниям</a></p>'
            f'<ol>{"".join(items)}</ol>'
            f'{appendix_section}{party_section}{match_section}'
            f'<details><summary>Оплата, реквизиты и ручная проверка</summary>'
            f'<pre style="white-space:pre-wrap">{details}</pre></details>'
            f'<p><a href="/api/reviews/{job_id}">Открыть JSON</a></p>')
    return page("Отчёт по договору", body)
