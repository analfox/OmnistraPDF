# Сборка одного exe: pyinstaller --noconfirm pdf_tools.spec  -> dist\OmnistraPDF.exe
a = Analysis(
    ['pdf_tools.py'],
    datas=[('ui', 'ui')],
    excludes=['numpy', 'onnxruntime', 'networkx', 'tkinter'],
)
# AI-разметка PyMuPDF (ONNX-модели, ~50 МБ) и заголовки для C-разработки приложению не нужны
skip = ('pymupdf\\layout', 'pymupdf\\mupdf-devel')
a.datas = [d for d in a.datas if not d[0].startswith(skip)]
a.binaries = [b for b in a.binaries if not b[0].startswith(skip)]

exe = EXE(
    PYZ(a.pure),
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='OmnistraPDF',
    icon='icon.ico',
    console=False,
)
