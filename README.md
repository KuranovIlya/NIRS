# NIRS
Перенос данных из Excel в Postgres<br>
Используемые библиотеки:<br>
    - psycopg2 - для работы с Postgres<br>
    - openpyxl - для четния Excel<br>
    - configparser - для парсинга файла инициализации<br>

Установка:<br>
pip install psycopg2<br>
pip install openpyxl<br>

Использован python 3.8

Вручную необходимо создать:
2 записи в таблице practice_kinds:
1. Учебная практика
2. Производственная практика

Запись в таблице opop_components, и указать, что к ней будут привязываться данные.

Записи в таблицы дисциплины под каждый вид практики.

Для каждого файла создавать новый учебный план.
