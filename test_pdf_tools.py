"""Самопроверка без окна: .venv\\Scripts\\python test_pdf_tools.py (для Word -> PDF нужен Word)."""
import functools
import hashlib
import http.server
import io
import json
import os
import shutil
import sys
import tempfile
import threading
import urllib.error

import pymupdf

import pdf_tools as t

sys.stdout.reconfigure(errors="replace")  # консоль cp1251 не умеет ✓/✗
t.os.startfile = lambda p: None
tmp = tempfile.mkdtemp(prefix="pdf тест ")
# путь длиннее 260 символов: Word такие сам не открывает
deep = os.path.join(tmp, *["Очень длинное название папки с лекциями по безопасности"] * 4)
os.makedirs(t.lp(deep))
assert len(deep) > 260


def make(name, pages=3, folder=tmp):
    p = os.path.join(folder, name)
    with pymupdf.open() as d:
        for i in range(pages):
            d.new_page().insert_text((72, 72), f"{name} {i + 1}")
        t.write_pdf(d, p)
    return p


def ok(r):
    assert r and "error" not in r, r
    return r


def picker(api, *paths):
    api._dialog = lambda *a, **k: list(paths)


def cropbox(doc, i):
    return [float(x) for x in doc.xref_get_key(doc[i].xref, "CropBox")[1].strip("[]").split()]


# разбор номеров, порядок имён, числа со словом
assert t.parse_ranges("1, 3-4,3", 5) == [(0, 0), (2, 3), (2, 2)]
for bad in ("0", "6", "2-1", "a", "-3"):
    try:
        t.parse_ranges(bad, 5)
        raise AssertionError(bad)
    except ValueError:
        pass
assert sorted(["10.pdf", "2.pdf", "1.pdf"], key=t.natural_key) == ["1.pdf", "2.pdf", "10.pdf"]
assert [t.plural(n, "файл", "файла", "файлов") for n in (1, 3, 5, 11, 22)] == \
       ["1 файл", "3 файла", "5 файлов", "11 файлов", "22 файла"]

# открыть по длинному пути, повернуть, удалить, отменить
api = t.Api()
a = make("a.pdf", folder=deep)
s = ok(api.open(a))
ids = [p["id"] for p in s["pages"]]
assert len(ids) == 3 and not s["dirty"]
s = ok(api.rotate([1], 90))
assert s["dirty"] and api._doc[1].rotation == 90
assert s["pages"][1]["id"] != ids[1] and s["pages"][0]["id"] == ids[0]  # миниатюра обновится только у второй
s = ok(api.delete([0]))
assert len(s["pages"]) == 2 and s["focus"] == 0
assert "error" in api.delete([0, 1])
ok(api.undo())
s = ok(api.undo())
assert [p["id"] for p in s["pages"]] == ids and not s["dirty"] and not s["canUndo"]

# картинка страницы; устаревший id — None
assert api.render(0, ids[0], 120, 160).startswith("data:image/jpeg;base64,")
assert api.render(0, -1, 120, 160) is None

# обрезка рамкой: срезать верхние 10 % видимой страницы
ok(api.rotate([0], 90))
ok(api.crop([0, 1], [0, 0.1, 1, 1]))
x0, y0, x1, y1 = cropbox(api._doc, 0)  # повёрнута на 90° по часовой: видимый верх — это левый край
assert abs(x0 - 59.5) < .01 and (y0, x1, y1) == (0, 595, 842), (x0, y0, x1, y1)
x0, y0, x1, y1 = cropbox(api._doc, 1)  # без поворота: верх страницы — это y1 в координатах PDF
assert (x0, y0, x1) == (0, 0, 595) and abs(y1 - (842 - 84.2)) < .01, (x0, y0, x1, y1)
assert "error" in api.crop([0], [0.5, 0, 0.5, 1])

# разделить
s = ok(api.split(""))
folder = os.path.join(deep, "a_части")
assert sorted(os.listdir(t.lp(folder))) == ["a_стр_1.pdf", "a_стр_2.pdf", "a_стр_3.pdf"], s
ok(api.split("1-2"))
with t.read_pdf(os.path.join(folder, "a_стр_1-2.pdf")) as d:
    assert d.page_count == 2
assert "error" in api.split("9")

# «Сохранить» — прямо в исходный файл по длинному пути, без окна
api._dialog = lambda *a, **k: (_ for _ in ()).throw(AssertionError("«Сохранить» не должно открывать окно"))
s = ok(api.save())
assert not s["dirty"] and s["name"] == "a.pdf"
with t.read_pdf(a) as d:
    assert d.page_count == 3 and d[0].rotation == 90

# «Сохранить как» — отдельный файл, дальше работа идёт с ним; оригинал не меняется
ok(api.rotate([1], 90))
copy = os.path.join(deep, "a копия.pdf")
picker(api, copy[:-4])  # без .pdf — допишется
s = ok(api.save_as())
assert s["name"] == "a копия.pdf" and not s["dirty"]
with t.read_pdf(a) as d:
    assert d[1].rotation == 0  # оригинал прежний
with t.read_pdf(copy) as d:
    assert d[1].rotation == 90
# в чужой существующий файл — сначала вопрос «Заменить?», запись только после save_to
other = make("другой.pdf", 1, folder=deep)
picker(api, other)
r = api.save_as()
assert r == {"overwrite": other, "name": "другой.pdf"}, r
with t.read_pdf(other) as d:
    assert d.page_count == 1
s = ok(api.save_to(other))
assert s["name"] == "другой.pdf"
with t.read_pdf(other) as d:
    assert d.page_count == 3

# закрытие с несохранённым: окно спрашивает, закрытие отменено
ok(api.rotate([0], 90))
asked = threading.Event()
api._window = type("W", (), {"run_js": lambda self, js: asked.set()})()
assert api._may_close() is False and asked.wait(2)
api._closing = True
assert api._may_close() is True

# объединить: новый документ в порядке имён, потом добавить в конец; битый файл документ не трогает
api = t.Api()
m2, m10 = make("2.pdf", 2), make("10.pdf", 1)
picker(api, m10, m2)
s = ok(api.merge())
assert s["dirty"] and s["isNew"] and s["name"] == "Объединённый.pdf"
assert [p.get_text().strip() for p in api._doc] == ["2.pdf 1", "2.pdf 2", "10.pdf 1"]
assert "error" in api.save()  # файла ещё нет — только «Сохранить как»
picker(api, m2)
s = ok(api.merge())
assert len(s["pages"]) == 5 and s["focus"] == 3
broken_pdf = os.path.join(tmp, "битый.pdf")
with open(broken_pdf, "w") as f:
    f.write("not a pdf")
picker(api, broken_pdf)
assert "error" in api.merge() and len(api._doc) == 5

# сжать: картинка 1500x1500 шума на 300pt (~360 dpi); текстовый файл сжать нечем
big = os.path.join(tmp, "big.pdf")
with pymupdf.open() as d:
    pix = pymupdf.Pixmap(pymupdf.csRGB, 1500, 1500, os.urandom(1500 * 1500 * 3), False)
    d.new_page().insert_image(pymupdf.Rect(72, 72, 372, 372), pixmap=pix)
    d.save(big)
api = t.Api()
ok(api.open(big))
s = ok(api.compress())
print(s["msg"])
assert s["dirty"] and "→" in s["msg"]
picker(api, os.path.join(tmp, "big_сжатый.pdf"))
ok(api.save_as())
assert os.path.getsize(os.path.join(tmp, "big_сжатый.pdf")) < os.path.getsize(big) / 5
api = t.Api()
ok(api.open(make("text.pdf")))
s = ok(api.compress())
assert s["kind"] == "info" and not s["dirty"] and not s["canUndo"]

# Word -> PDF: файл с путём длиннее 259 символов (Word такие сам не открывает) и битый файл
import win32com.client

w = win32com.client.DispatchEx("Word.Application")
short = os.path.join(tmp, "w.docx")
doc = w.Documents.Add()
doc.Content.Text = "Привет"
doc.SaveAs2(short, 16)
doc.Close(0)
w.Quit(0)
long_docx = os.path.join(deep, "Документ с длинным названием.docx")
shutil.copyfile(short, t.lp(long_docx))
broken = os.path.join(tmp, "битый.docx")
with open(broken, "w") as f:
    f.write("not a docx")
api = t.Api()
picker(api, long_docx, broken)
s = ok(api.word())
print(s["msg"])
assert s["open"] and s["kind"] == "warning" and "✓ Документ" in s["msg"] and "✗ битый.docx" in s["msg"]
assert "Привет" in api._doc[0].get_text()
assert os.path.exists(t.lp(long_docx[:-5] + ".pdf"))

# обновления: сравнение версий
assert t.parse_version("v1.10.0") > t.parse_version("1.9.3") > t.parse_version("v1.9")
real_urlopen = t.urllib.request.urlopen


def offline(*a, **k):
    raise urllib.error.URLError("нет сети")


# нет интернета: при запуске молча None, по кнопке — понятная ошибка
t.urllib.request.urlopen = offline
assert t.Api().check_update() is None
assert "Нет связи" in t.Api().update()["error"]

# новая версия на «GitHub»: exe раздаёт локальный сервер
srv_dir = os.path.join(tmp, "srv")
os.makedirs(srv_dir)
new_exe = b"MZ" + os.urandom(5000)
with open(os.path.join(srv_dir, "new.exe"), "wb") as f:
    f.write(new_exe)
quiet = type("Quiet", (http.server.SimpleHTTPRequestHandler,), {"log_message": lambda *a: None})
srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), functools.partial(quiet, directory=srv_dir))
threading.Thread(target=srv.serve_forever, daemon=True).start()
asset = {"name": "OmnistraPDF.exe", "browser_download_url": f"http://127.0.0.1:{srv.server_port}/new.exe",
         "size": len(new_exe), "digest": "sha256:" + hashlib.sha256(new_exe).hexdigest()}
release = {"tag_name": "v9.0.0", "html_url": "https://example.invalid", "assets": [asset]}


def fake_urlopen(req, timeout=None):
    if "api.github.com" in req.full_url:
        return io.BytesIO(json.dumps(release).encode())
    return real_urlopen(req, timeout=timeout)


t.urllib.request.urlopen = fake_urlopen
assert t.Api().check_update() == {"version": "9.0.0"}
assert t.Api().update(False) == {"version": "9.0.0"}
opened = []
t.webbrowser.open = opened.append  # из исходников себя не подменяем — открываем страницу релиза
assert "в браузере" in t.Api().update(True)["msg"] and opened == ["https://example.invalid"]

exe = os.path.join(tmp, "OmnistraPDF.exe")
with open(exe, "wb") as f:
    f.write(b"old")
t.install_update(asset, exe)
assert open(exe, "rb").read() == new_exe and open(exe + ".old", "rb").read() == b"old"
try:  # скачалось не то: exe не трогаем
    t.install_update(dict(asset, digest="sha256:" + "0" * 64), exe)
    raise AssertionError("битое обновление установилось")
except RuntimeError:
    pass
assert open(exe, "rb").read() == new_exe and not os.path.exists(exe + ".new")

release["tag_name"] = "v" + t.VERSION
assert t.Api().check_update() is None
assert "последняя" in t.Api().update()["msg"]

# тег есть, а exe ещё собирается — не «последняя версия»
release.update(tag_name="v9.0.0", assets=[])
assert t.Api().check_update() is None
assert "собирается" in t.Api().update()["msg"]


def no_releases(req, timeout=None):
    raise urllib.error.HTTPError(req.full_url, 404, "Not Found", {}, None)


t.urllib.request.urlopen = no_releases  # релизов нет совсем
assert t.Api().check_update() is None
assert "нет готовой версии" in t.Api().update()["msg"]
t.urllib.request.urlopen = real_urlopen
srv.shutdown()

shutil.rmtree(t.lp(tmp))
print("OK")
