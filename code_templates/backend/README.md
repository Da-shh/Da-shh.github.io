# tropa-api — серверная часть ИС «Тропа»

Веб-сервис построения туристических маршрутов по Тульской области для людей с
ограниченными возможностями здоровья. Серверная часть.

## Стек

- Python 3.12, FastAPI (REST API, JSON)
- PostgreSQL 16 + PostGIS (геоданные), SQLAlchemy
- Redis (кеширование)
- YandexGPT (автоматизированный сбор сведений о доступности)
- Docker

## Быстрый старт (Docker)

```bash
cp .env.example .env   # заполните значения
docker compose up --build
```

API будет доступно на `http://localhost:8000`, документация — `http://localhost:8000/docs`.

## Запуск вручную (без Docker)

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt  # или: pip install -e .
cp .env.example .env             # заполните значения
alembic upgrade head             # применить миграции БД
uvicorn app.main:app --reload    # запустить сервер
```

> Требуется поднятый PostgreSQL с расширением PostGIS и Redis (локально или в Docker).

## Переменные окружения

См. `.env.example`. Реальный файл `.env` хранится только локально и в репозиторий
не коммитится.

## Важно по безопасности

Никогда не коммитьте `.env`, ключи (`*.key`, `key.json`), дампы БД (`*.sql`, `*.dump`) —
они исключены в `.gitignore`. В дампах могут быть персональные данные пользователей
(152-ФЗ).
