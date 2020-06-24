# NIRS
Перенос данных из Excel в Postgres
Используемые библиотеки:
    - psycopg2 - для работы с Postgres
    - openpyxl - для четния Excel
    - configparser - для парсинга файла инициализации

Установка:
pip install psycopg2
pip install openpyxl

Использован python 3.8

Вручную необходимо создать:
2 записи в таблице practice_kinds:
1. Учебная практика
2. Производственная практика

Запись в таблице opop_components, и указать, что к ней будут привязываться данные.
