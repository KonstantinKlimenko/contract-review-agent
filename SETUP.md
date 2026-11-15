# Локальный запуск

Инструкция рассчитана на Windows с Docker Desktop, Python 3.11 и Ollama. Команды выполняются из корня репозитория. Схема использует локальные порты 5678 для n8n, 8765 для API и 11434 для Ollama.

## Подготовить API и модели

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\hf.exe download BAAI/bge-m3 --local-dir models/bge-m3
ollama pull qwen3.5:4b
.\.venv\Scripts\python.exe -B -m uvicorn contract_review_api:app --host 127.0.0.1 --port 8765
```

Ollama должен быть запущен на хосте. Путь к уже скачанному BGE-M3 можно задать переменной `EMBEDDING_MODEL_PATH`. API создаёт SQLite в каталоге `data`. Для другого каталога используйте `CONTRACT_REVIEW_DATA`.

## Импортировать и настроить n8n

1. В другом окне PowerShell запустите n8n.

   ```powershell
   docker compose up -d
   ```

   Зарегистрируйте локального владельца по адресу http://localhost:5678, затем импортируйте схемы.

   ```powershell
   docker compose exec -T n8n n8n import:workflow --input=/workflows --separate
   ```

2. В дочерних workflow **«Договоры — агент одиночной проверки»** и **«Договоры — агент пакетной проверки»** откройте узел **Ollama Chat Model** и привяжите локальные учётные данные Ollama с базовым адресом `http://host.docker.internal:11434`.
3. В основном workflow **«Договоры — учебный AI Agent»** откройте webhook **«Вопрос из чата по договору»**. Создайте учётные данные **Header Auth** с именем заголовка `X-Contract-Chat-Token`. Скопируйте значение из локального API командой ниже и вставьте из буфера обмена в поле **Value**. Базу и значение заголовка храните локально.

   ```powershell
   .\.venv\Scripts\python.exe -B -c "from contract_review_api import chat_bridge_token; print(chat_bridge_token(), end='')" | Set-Clipboard
   ```

4. Сохраните изменения и опубликуйте workflow в n8n. Для основной формы опубликуйте **«Договоры — учебный AI Agent»** после дочерних workflow. При необходимости выберите импортированный дочерний workflow заново в узлах вызова. В экспортированных схемах сохранены исходные идентификаторы связей.

## Повторить демонстрацию

1. Откройте [загрузку шаблонов](http://localhost:5678/form/contract-review-templates) и активируйте `examples/company-template-v1.docx` и `examples/approved-demo.xlsx`.
2. В [форме пакетной проверки](http://localhost:5678/form/contract-agent-batch) выберите шаблон и приложите `examples/03-terms-portfolio-demo.pdf`, `examples/04-missing-appendix.docx` и `examples/carrier-demo.xlsx`.
3. В сводке спросите, где изменился срок оплаты, какого приложения не хватает и на сколько вырос тариф. В примере источники находятся в п. 5.2 на стр. 6 PDF, п. 1.2 DOCX и ячейке `Тарифы!E2` двух книг.

Локальные тесты без запуска n8n и моделей:

```powershell
.\.venv\Scripts\python.exe -B -m unittest discover -p "test_compare_contract*.py"
```
