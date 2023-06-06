# Краткое описание

Сервис собирает пожертвования на различные целевые проекты: лечение, еду, кров и другую помощь хвостатым друзьям. Так же можно сделать таблицу со всеми закрытыми проектами, отсортированными по времени сборов. 

# Порядок запуска приложения 

Клонировать репозиторий и перейти в него в командной строке:

```
git clone 
```

```
cd cat_charity_fund
```

Cоздать и активировать виртуальное окружение:

```
python3 -m venv venv
```

* Если у вас Linux/macOS

    ```
    source venv/bin/activate
    ```

* Если у вас windows

    ```
    source venv/scripts/activate
    ```

Установить зависимости из файла requirements.txt:

```
python3 -m pip install --upgrade pip
```

```
pip install -r requirements.txt
```

Создать базу данных и применить миграции:

```
alembic upgrade head
```
Важное замечание: в файле app/core/config.py задаются переменные окружения и их дефолтные значения. Перезаписать свои можно в файле .env, разместив его в корневой директории. Переменные лучше полностью прописывать заглавными в файле .env, а в файле config.py строчными в соответствии с PEP8:

```
# config.py
database_url: str = 'sqlite+aiosqlite:///./fastapi.db'
# .env
DATABASE_URL=sqlite+aiosqlite:///./fastapi.db
```

Запуск приложения:

```
uvicorn app.main:app
```

По адресу /docs/ находится документация проекта в формате swagger


# Список использованных технологий

- Python (язык разработки)
- FastAPI (фрэймворк приложения)
- Uvicorn (ASGI-сервер)
- SQLAlchemy (для работы с базами данных)
- Alembic (для создания миграций)
- Pydantic (для создания схем)
- Swagger (документация проекта)
- Google Drive API (api для работы с диском)
- Google Sheets API (api для работы с таблицами)

# Автор

[Андрей А.](https://github.com/Ander-dog)