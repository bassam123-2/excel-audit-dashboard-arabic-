# -*- mode: python ; coding: utf-8 -*-
# One-file build → dist/excel_arabic_dashboard.exe
# Build: pyinstaller excel_arabic_dashboard.spec

from PyInstaller.utils.hooks import collect_submodules

flask_hidden = collect_submodules("flask")
jinja_hidden = collect_submodules("jinja2")
django_hidden = collect_submodules("django")
rest_hidden = collect_submodules("rest_framework")

a = Analysis(
    ["excel_arabic_desktop.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("templates", "templates"),
        ("assets", "assets"),
        ("django_project", "django_project"),
        ("apps", "apps"),
    ],
    hiddenimports=[
        *flask_hidden,
        *jinja_hidden,
        *django_hidden,
        *rest_hidden,
        "pandas",
        "openpyxl",
        "pptx",
        "pymysql",
        "werkzeug",
        "werkzeug.middleware.proxy_fix",
        "email",
        "email.mime",
        "email.mime.multipart",
        "email.mime.text",
        "email.mime.image",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="excel_arabic_dashboard",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
