Kumex Stock Calc
================

Локальная настройка на любой машине
-----------------------------------
1) Убедись, что установлен Python 3.12+.
2) В корне проекта создай окружение:
   ```
   python -m venv .venv
   ```
3) Установи зависимости:
   ```
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```
4) Запуск приложения:
   ```
   .\.venv\Scripts\python.exe "src\kumex\Kumex Ladu.py"
   ```
   (Если PowerShell блокирует `Activate.ps1`, не активируй сессию — просто вызывай python из `.venv\Scripts` как выше.)

Замечания
---------
- `.venv` хранится только локально и игнорируется git/OneDrive (см. `.gitignore`).
- Для каждой машины окружение создаётся заново, код и зависимости берутся из репозитория/OneDrive.
