"""OmnistraPDF: просмотр и правка PDF, Word -> PDF.

Окно — HTML в ui/index.html (WebView2), вся работа с файлами — здесь, в классе Api.
Исходный файл не меняется, пока не нажали «Сохранить».
Обновления — релизы GitHub: тег vX.Y.Z и файл OmnistraPDF.exe (собирает .github/workflows/release.yml).
"""
import base64
import ctypes
import functools
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import webbrowser

import pymupdf
import webview

VERSION = "1.0.4"  # = тег релиза без «v»; поднимает «залить на гитхаб.bat»
RELEASES = "https://github.com/analfox/OmnistraPDF/releases"
PDF_TYPES = ("PDF (*.pdf)",)
WORD_TYPES = ("Word (*.doc;*.docx;*.docm;*.rtf;*.odt)", "Все файлы (*.*)")
# ponytail: для отмены документ целиком хранится в памяти; на сканах в сотни МБ стоит хранить снимки на диске
UNDO_LIMIT = 15


def lp(path):
    """Путь с префиксом \\\\?\\ — так Windows открывает пути длиннее 260 символов."""
    path = os.path.abspath(path)
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):  # сетевая папка \\server\share
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def read_pdf(path):
    with open(lp(path), "rb") as f:
        doc = pymupdf.open(stream=f.read(), filetype="pdf")
    if doc.needs_pass:
        raise ValueError(f"«{os.path.basename(path)}» защищён паролем.")
    return doc


def pdf_bytes(doc):
    return doc.tobytes(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, use_objstms=1)


def write_pdf(doc, path):
    """Пишет через временный файл: сбой посреди записи не испортит уже лежащий PDF."""
    data = pdf_bytes(doc)
    tmp = lp(path) + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, lp(path))


def plural(n, one, few, many):
    if n % 10 == 1 and n % 100 != 11:
        return f"{n} {one}"
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return f"{n} {few}"
    return f"{n} {many}"


def mb(size):
    return f"{size / 1024 / 1024:.1f} МБ".replace(".", ",")


def parse_ranges(text, count):
    """'1,3,5-7' -> [(0,0), (2,2), (4,6)]; в тексте нумерация с 1."""
    ranges = []
    for part in re.split(r"[,;\s]+", (text or "").strip()):
        if not part:
            continue
        a, _, b = part.partition("-")
        try:
            a, b = int(a), int(b or a)
        except ValueError:
            raise ValueError(f"Не понимаю «{part}». Нужны номера страниц через запятую: 1-3, 4-10") from None
        if not 1 <= a <= b <= count:
            raise ValueError(f"Нет страниц «{part}»: в документе {plural(count, 'страница', 'страницы', 'страниц')}.")
        ranges.append((a - 1, b - 1))
    return ranges


def natural_key(path):
    """Порядок имён как в проводнике: «2.pdf» раньше «10.pdf»."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", os.path.basename(path))]


def crop_rect(page, top, right, bottom, left):
    """Поля в пунктах — как страницу видно на экране. Возвращает новый CropBox: он задаётся без учёта поворота."""
    m = (top, right, bottom, left)
    k = page.rotation // 90
    top, right, bottom, left = (m[(j + k) % 4] for j in range(4))
    c = page.cropbox
    r = pymupdf.Rect(c.x0 + left, c.y0 + top, c.x1 - right, c.y1 - bottom)
    if r.is_empty or r.width < 1 or r.height < 1:
        raise ValueError(f"На странице {page.number + 1} от области обрезки ничего не остаётся.")
    return r


def com_message(e):
    info = getattr(e, "excepinfo", None)  # у ошибок COM текст Word лежит тут
    return info[2].split("\n")[0].strip() if info and info[2] else str(e)


def human(e):
    if isinstance(e, PermissionError):
        return "Файл занят другой программой или нет прав на запись. Закройте его и повторите."
    if isinstance(e, FileNotFoundError):
        return "Файл не найден: возможно, его переместили или переименовали."
    if isinstance(e, pymupdf.FileDataError):
        return "Не получилось прочитать PDF: файл повреждён или это не PDF."
    if isinstance(getattr(e, "Message", None), str):  # ошибка .NET: str() добавил бы стек вызовов
        return e.Message
    return com_message(e) or e.__class__.__name__


def word_to_pdf(files):
    """Переводит Word-файлы в PDF рядом с ними. Возвращает (строки отчёта, пути готовых PDF)."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()  # вызов приходит из потока окна, а не из главного
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception:
            raise RuntimeError("Не найден Microsoft Word: он нужен, чтобы перевести файл в PDF.") from None
        word.DisplayAlerts = 0
        lines, done = [], []
        try:
            for f in files:
                pdf = os.path.splitext(f)[0] + ".pdf"
                try:
                    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
                        # Word не открывает пути длиннее 259 символов — работаем с копией в короткой папке
                        stem = os.path.splitext(os.path.basename(f))[0][:100]
                        src = os.path.join(tmp, stem + os.path.splitext(f)[1])
                        out = os.path.join(tmp, stem + ".pdf")
                        shutil.copyfile(lp(f), src)
                        try:
                            # фиктивный пароль: защищённый файл даст ошибку, а не окно пароля в невидимом Word
                            doc = word.Documents.Open(src, False, True, False, "x")
                        except Exception as e:
                            raise RuntimeError(com_message(e)) from None
                        if doc is None:
                            raise RuntimeError("Word не смог открыть файл.")
                        try:
                            doc.ExportAsFixedFormat(out, 17)  # 17 = PDF
                        finally:
                            doc.Close(0)
                        shutil.copyfile(out, lp(pdf))
                    done.append(pdf)
                    lines.append(f"✓ {os.path.basename(f)}")
                except Exception as e:
                    lines.append(f"✗ {os.path.basename(f)}: {human(e)}")
        finally:
            word.Quit(0)
        return lines, done
    finally:
        pythoncom.CoUninitialize()


def parse_version(tag):
    return tuple(int(x) for x in re.findall(r"\d+", tag)[:3])


def gh_open(url, timeout, method="GET"):
    req = urllib.request.Request(url, method=method, headers={"User-Agent": f"OmnistraPDF/{VERSION}"})
    return urllib.request.urlopen(req, timeout=timeout)


def latest_version(timeout):
    """Номер последнего релиза. Не через API GitHub: там лимит 60 запросов в час на IP, а в офисе за одним
    адресом его выбирают за утро (ответ 403). /releases/latest перенаправляет на /releases/tag/vX.Y.Z;
    релизов нет — на /releases, тогда None."""
    with gh_open(f"{RELEASES}/latest", timeout, "HEAD") as r:
        m = re.search(r"/tag/v?([^/?#]+)$", r.geturl())
    return m.group(1) if m else None


def release_sha256(version, timeout):
    """Контрольная сумма exe, её выкладывает сборка вместе с exe. Нет файла (HTTP 404) — релиз ещё не готов."""
    with gh_open(f"{RELEASES}/download/v{version}/OmnistraPDF.exe.sha256", timeout) as r:
        return r.read(200).decode("ascii").split()[0].lower()


def install_update(version, sha256, exe):
    """Скачивает новый exe и ставит на место exe. Запущенный exe Windows не даёт перезаписать, но даёт переименовать."""
    new, old = exe + ".new", exe + ".old"
    h = hashlib.sha256()
    with gh_open(f"{RELEASES}/download/v{version}/OmnistraPDF.exe", 60) as r, open(new, "wb") as f:
        while chunk := r.read(1 << 20):
            f.write(chunk)
            h.update(chunk)
    if h.hexdigest() != sha256:
        os.remove(new)
        raise RuntimeError("Обновление скачалось с ошибкой. Попробуйте ещё раз.")
    try:
        os.remove(old)  # остался от прошлого обновления
    except OSError:
        pass
    os.replace(exe, old)
    try:
        os.replace(new, exe)
    except OSError:
        os.replace(old, exe)
        raise


def file_dialog(uid, save=False, multiple=False, types=PDF_TYPES, directory="", name=""):
    """Окно «Открыть» / «Сохранить как» — то же, что у pywebview, но без проверок .NET Framework: для путей
    длиннее 260 символов он отвечает «файл не существует» и не даёт выбрать файл. Есть ли файл и не затрём ли
    другой — проверяем сами (lp, «Заменить?» в окне программы)."""
    from webview.platforms import winforms as wf  # окно уже открыто, модуль загружен

    d = wf.WinForms.SaveFileDialog() if save else wf.WinForms.OpenFileDialog()
    if save:
        d.OverwritePrompt = False
        d.FileName = name
    else:
        d.Multiselect = multiple
        d.CheckFileExists = False
    d.CheckPathExists = False
    d.ValidateNames = False
    d.Filter = "|".join("{0} ({1})|{1}".format(*wf.parse_file_type(t)) for t in types)
    d.InitialDirectory = directory if directory and os.path.isdir(lp(directory)) else os.path.expanduser("~")
    d.RestoreDirectory = True
    if d.ShowDialog(wf.BrowserView.instances[uid]) != wf.WinForms.DialogResult.OK:
        return []
    # не d.FileNames: этот геттер прогоняет путь через Path.GetFullPath и падает на длинных путях
    import clr
    from System.Reflection import BindingFlags
    field = clr.GetClrType(wf.WinForms.FileDialog).GetField("fileNames", BindingFlags.NonPublic | BindingFlags.Instance)
    return list(field.GetValue(d) or [])


def api_call(fn):
    """Методы окна идут по одному (PyMuPDF не потокобезопасен), ошибка уходит в окно текстом."""
    @functools.wraps(fn)
    def wrapper(self, *args):
        try:
            with self._lock:
                return fn(self, *args)
        except Exception as e:
            return {"error": human(e)}
    return wrapper


class Api:
    """Всё, что вызывает окно (window.pywebview.api.*). Методы с «_» окну не видны."""

    def __init__(self):
        self._lock = threading.Lock()
        self._window = None
        self._doc = None
        self._path = None
        self._undo = []    # снимки (байты PDF, id страниц) для «Отменить»
        self._saved = 0    # длина _undo в момент сохранения: не совпадает — есть несохранённое
        self._new = False  # документа ещё нет на диске (склеен из нескольких)
        self._ids = []     # id страниц: окно хранит по ним миниатюры и перерисовывает только изменённые
        self._next_id = 0
        self._closing = False

    # ---------- служебное ----------
    def _new_ids(self, n):
        self._next_id += n
        return list(range(self._next_id - n, self._next_id))

    def _load(self, doc, path, saved=True):
        """saved=False — документа ещё нет на диске (склеен из нескольких): «Сохранить» спросит, куда."""
        self._doc, self._path, self._undo = doc, path, []
        self._saved = 0 if saved else -1
        self._new = not saved
        self._ids = self._new_ids(len(doc))

    def _dirty(self):
        return self._doc is not None and len(self._undo) != self._saved

    def _snapshot(self):
        self._undo.append((self._doc.tobytes(), list(self._ids)))
        extra = len(self._undo) - UNDO_LIMIT
        if extra > 0:
            del self._undo[:extra]
            self._saved -= extra

    def _restore(self):
        data, self._ids = self._undo.pop()
        self._doc = pymupdf.open(stream=data, filetype="pdf")

    def _touch(self, pages):
        for i in pages:
            self._ids[i] = self._new_ids(1)[0]

    def _state(self, focus=None, msg=None, kind="success"):
        s = {"open": self._doc is not None, "version": VERSION}
        if self._doc is not None:
            s.update(name=os.path.basename(self._path), path=self._path, dirty=self._dirty(),
                     canUndo=bool(self._undo), isNew=self._new,
                     pages=[{"id": i, "w": p.rect.width, "h": p.rect.height} for i, p in zip(self._ids, self._doc)])
        if focus is not None:
            s["focus"] = focus
        if msg:
            s["msg"], s["kind"] = msg, kind
        return s

    def _dialog(self, kind, multiple=False, types=PDF_TYPES, directory="", name=""):
        return file_dialog(self._window.uid, kind == webview.FileDialog.SAVE, multiple, types, directory, name)

    def _write(self, path):
        write_pdf(self._doc, path)
        self._path, self._saved, self._new = path, len(self._undo), False
        return self._state()

    def _may_close(self):
        if self._closing or not self._dirty():
            return True
        # спросить «Сохранить?» окном страницы; из обработчика закрытия напрямую нельзя — он в потоке окна
        threading.Thread(target=self._window.run_js, args=("app.askClose()",), daemon=True).start()
        return False

    # ---------- вызовы из окна ----------
    @api_call
    def state(self):
        return self._state()

    @api_call
    def open(self, path=None):
        if path is None:
            paths = self._dialog(webview.FileDialog.OPEN)
            if not paths:
                return None
            path = paths[0]
        self._load(read_pdf(path), path)
        return self._state(focus=0)

    @api_call
    def word(self):
        files = self._dialog(webview.FileDialog.OPEN, multiple=True, types=WORD_TYPES)
        if not files:
            return None
        lines, done = word_to_pdf(files)
        if done:
            self._load(read_pdf(done[0]), done[0])
        failed = len(files) - len(done)
        msg = None
        if failed or len(files) > 1:
            msg = f"В PDF переведено {len(done)} из {len(files)}. PDF лежат рядом с исходными файлами.\n" + "\n".join(lines)
        return self._state(focus=0 if done else None, msg=msg, kind="warning" if failed else "success")

    @api_call
    def merge(self):
        files = self._dialog(webview.FileDialog.OPEN, multiple=True)
        if not files:
            return None
        files.sort(key=natural_key)
        sources = [read_pdf(f) for f in files]  # сначала прочитать все: битый файл не оставит документ склеенным наполовину
        new = self._doc is None
        if new:
            doc, start = pymupdf.open(), 0
        else:
            self._snapshot()
            doc, start = self._doc, len(self._doc)
        for src in sources:
            doc.insert_pdf(src)
        if new:
            self._load(doc, os.path.join(os.path.dirname(files[0]), "Объединённый.pdf"), saved=False)
        else:
            self._ids += self._new_ids(len(doc) - start)
        names = "\n".join(os.path.basename(f) for f in files)
        return self._state(focus=start, msg=f"Добавлено по порядку имён:\n{names}")

    @api_call
    def rotate(self, pages, angle):
        self._snapshot()
        for i in pages:
            page = self._doc[i]
            page.set_rotation((page.rotation + angle) % 360)
        self._touch(pages)
        return self._state()

    @api_call
    def delete(self, pages):
        pages = sorted(set(pages))
        if len(pages) >= len(self._doc):
            raise ValueError("Все страницы удалить нельзя: в документе должна остаться хотя бы одна.")
        self._snapshot()
        self._doc.delete_pages(pages)
        for i in reversed(pages):
            del self._ids[i]
        return self._state(focus=min(pages[0], len(self._doc) - 1))

    @api_call
    def crop(self, pages, box):
        """box — что оставить, доли видимой страницы: [слева, сверху, справа, снизу]."""
        left, top, right, bottom = box
        rects = []
        for i in pages:  # сначала проверить все страницы, потом менять
            page = self._doc[i]
            w, h = page.rect.width, page.rect.height
            rects.append(crop_rect(page, top * h, (1 - right) * w, (1 - bottom) * h, left * w))
        self._snapshot()
        for i, r in zip(pages, rects):
            self._doc[i].set_cropbox(r)
        self._touch(pages)
        return self._state()

    @api_call
    def compress(self):
        before = len(pdf_bytes(self._doc))
        self._snapshot()
        self._doc.rewrite_images(dpi_threshold=160, dpi_target=120, quality=60)
        after = len(pdf_bytes(self._doc))
        if after > before * 0.95:
            self._restore()
            return self._state(msg="Сжать заметно не получилось: крупных картинок в файле нет.", kind="info")
        self._touch(range(len(self._doc)))
        return self._state(msg=f"Файл сжат: {mb(before)} → {mb(after)}. Сохраните его, чтобы записать изменения.")

    @api_call
    def split(self, text):
        n = len(self._doc)
        ranges = parse_ranges(text, n) or [(i, i) for i in range(n)]
        base = os.path.splitext(self._path)[0]
        folder, stem = base + "_части", os.path.basename(base)
        os.makedirs(lp(folder), exist_ok=True)
        for a, b in ranges:
            with pymupdf.open() as part:
                part.insert_pdf(self._doc, from_page=a, to_page=b)
                label = f"{a + 1}" if a == b else f"{a + 1}-{b + 1}"
                write_pdf(part, os.path.join(folder, f"{stem}_стр_{label}.pdf"))
        try:
            os.startfile(folder)
        except OSError:
            pass  # проводник не открывает слишком длинные пути — сообщение всё равно скажет, где файлы
        files = plural(len(ranges), "файл", "файла", "файлов")
        return self._state(msg=f"Документ разделён на {files} в папке «{os.path.basename(folder)}».")

    @api_call
    def undo(self):
        if self._undo:
            self._restore()
        return self._state()

    @api_call
    def save(self):
        """В тот же файл. Склеенный документ ещё не на диске — окно отправит его в «Сохранить как»."""
        if self._new:
            raise ValueError("Документ ещё не сохранён в файл — выберите «Сохранить как».")
        return self._write(self._path)

    @api_call
    def save_as(self):
        paths = self._dialog(webview.FileDialog.SAVE, directory=os.path.dirname(self._path),
                             name=os.path.basename(self._path))
        if not paths:
            return None
        path = paths[0] if paths[0].lower().endswith(".pdf") else paths[0] + ".pdf"
        if os.path.exists(lp(path)) and os.path.normcase(path) != os.path.normcase(self._path):
            return {"overwrite": path, "name": os.path.basename(path)}  # окно спросит «Заменить?» и вызовет save_to
        return self._write(path)

    @api_call
    def save_to(self, path):
        """Запись после «Заменить?» в окне программы."""
        return self._write(path)

    @api_call
    def render(self, i, page_id, width, height):
        """Картинка страницы, вписанная в width×height пикселей. None — страница уже другая."""
        if self._doc is None or i >= len(self._ids) or self._ids[i] != page_id:
            return None
        page = self._doc[i]
        zoom = min(width / page.rect.width, height / page.rect.height)
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        return "data:image/jpeg;base64," + base64.b64encode(pix.tobytes("jpg", jpg_quality=90)).decode()

    def check_update(self):
        """Тихая проверка при запуске: нет интернета или GitHub не ответил — None, проверим при следующем запуске."""
        try:
            new = latest_version(timeout=5)
            if new and parse_version(new) > parse_version(VERSION):
                release_sha256(new, timeout=5)  # exe уже выложен
                return {"version": new}
        except Exception:
            pass
        return None

    @api_call
    def update(self, install=False):
        """install=False — только узнать, есть ли новая версия; окно спросит про несохранённое и вызовет с True."""
        try:
            new = latest_version(timeout=10)
            if not new:
                return self._state(msg="На GitHub пока нет готовой версии для скачивания.", kind="info")
            if parse_version(new) <= parse_version(VERSION):
                return self._state(msg=f"Установлена последняя версия: {VERSION}.")
            try:
                sha = release_sha256(new, timeout=10)
            except urllib.error.HTTPError as e:
                if e.code != 404:
                    raise
                return self._state(msg=f"Версия {new} ещё собирается на GitHub. Попробуйте через несколько минут.",
                                   kind="info")
        except urllib.error.HTTPError as e:
            raise RuntimeError(f"GitHub ответил ошибкой {e.code}. Попробуйте позже.") from None
        except OSError:
            raise RuntimeError("Нет связи с GitHub. Проверьте интернет и повторите.") from None
        if not install:
            return {"version": new}
        page = f"{RELEASES}/tag/v{new}"
        if not getattr(sys, "frozen", False):  # запуск из исходников: себя не подменяем
            webbrowser.open(page)
            return self._state(msg=f"Версия {new} открыта в браузере.", kind="info")
        try:
            install_update(new, sha, sys.executable)
        except PermissionError:
            webbrowser.open(page)
            raise RuntimeError("Нет прав на запись в папку программы. "
                               "Страница новой версии открыта в браузере: скачайте её вручную.") from None
        reopen = [self._path] if self._doc is not None and not self._dirty() else []
        # PYINSTALLER_RESET_ENVIRONMENT: новый exe — самостоятельная программа, а не дочерний процесс этой
        subprocess.Popen([sys.executable, *reopen], env={**os.environ, "PYINSTALLER_RESET_ENVIRONMENT": "1"})
        self._closing = True
        threading.Timer(0.5, self._window.destroy).start()  # сначала ответ окну, потом закрыть
        return {"restarting": True}

    def quit(self):
        self._closing = True
        self._window.destroy()


def remove_leftovers(exe):
    """Старый exe после обновления. Только удаляет файл — программу не трогает. Одна попытка: прошлая копия
    сначала закрывается (убирает свою временную папку) и держит файл, поэтому ждём 15 с. Не вышло —
    удалится при следующем запуске."""
    time.sleep(15)
    for path in (exe + ".old", exe + ".new"):
        try:
            os.remove(path)
        except OSError:
            pass


def main():
    if getattr(sys, "frozen", False):
        threading.Thread(target=remove_leftovers, args=(sys.executable,), daemon=True).start()
    api = Api()
    if len(sys.argv) > 1:  # файл, перетащенный на exe или открытый через «Открыть с помощью»
        try:
            api._load(read_pdf(sys.argv[1]), sys.argv[1])
        except Exception:
            pass
    here = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    window = webview.create_window("OmnistraPDF", os.path.join(here, "ui", "index.html"), js_api=api,
                                   width=1280, height=820, min_size=(980, 620), maximized=True,
                                   background_color="#f6f5fa")
    api._window = window
    window.events.closing += api._may_close
    try:
        webview.start(gui="edgechromium")
    except Exception:
        ctypes.windll.user32.MessageBoxW(
            0, "Для окна программы нужен Microsoft Edge WebView2 Runtime. Установите его с сайта Microsoft "
               "(бесплатно) и запустите программу снова.", "OmnistraPDF", 0x10)


if __name__ == "__main__":
    main()
