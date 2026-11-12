# OmnistraPDF

Программа для Windows: просмотр PDF, перевод Word в PDF, объединение и разделение файлов, поворот, удаление и обрезка страниц, сжатие.

## Скачать

[**OmnistraPDF.exe**](https://github.com/analfox/OmnistraPDF/releases/latest/download/OmnistraPDF.exe) — один файл, устанавливать не нужно. При запуске программа сама проверяет, вышла ли новая версия; без интернета просто работает дальше.

Нужно:
- Windows 10 или 11 с WebView2 — в Windows 11 он есть, в Windows 10 ставится вместе с Edge;
- Microsoft Word — только для кнопки «Word в PDF».

Windows может предупредить, что программа скачана из интернета: «Подробнее» → «Выполнить в любом случае».

## Разработка

```
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python pdf_tools.py              # запустить
.venv\Scripts\python test_pdf_tools.py         # самопроверка (для Word -> PDF нужен Word)
.venv\Scripts\pyinstaller --noconfirm pdf_tools.spec   # dist\OmnistraPDF.exe
```

Окно — `ui/index.html` (WebView2), работа с PDF — `pdf_tools.py` (PyMuPDF), Word — через установленный Word.

## Новая версия

1. Поднять `VERSION` в `pdf_tools.py` и закоммитить.
2. `git tag v1.0.1` и `git push origin main v1.0.1`.

GitHub Actions соберёт exe и выложит релиз; программы у всех при следующем запуске предложат обновиться.
Если минут через 10 релиза нет: вкладка Actions → «Релиз» → Run workflow → в списке выбрать тег версии.

Шрифт Onest — SIL Open Font License, `ui/OFL.txt`.
