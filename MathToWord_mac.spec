# -*- mode: python ; coding: utf-8 -*-
# PyInstaller-спецификация для сборки MathToWord.app под macOS (запуск: pyinstaller --noconfirm MathToWord_mac.spec).
# Файл `pandoc` (macOS-бинарник нужной архитектуры) должен лежать рядом с этим .spec — его кладёт workflow.

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[('pandoc', '.')],   # вшиваем pandoc в корень пакета (sys._MEIPASS)
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='MathToWord',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,   # для Tk должно быть False
    target_arch=None,       # архитектура = архитектура runner'а (arm64 или x86_64)
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='MathToWord',
)

app = BUNDLE(
    coll,
    name='MathToWord.app',
    icon=None,
    bundle_identifier='com.mathtoword.app',
    info_plist={
        'CFBundleName': 'MathToWord',
        'CFBundleDisplayName': 'MathToWord',
        'CFBundleShortVersionString': '1.0.0',
        'LSMinimumSystemVersion': '11.0',
        'NSHighResolutionCapable': True,
        # Без этой строки macOS молча запрещает управлять Word через AppleScript
        'NSAppleEventsUsageDescription':
            'MathToWord управляет Microsoft Word, чтобы скопировать текст с формулами в буфер обмена.',
        'NSDesktopFolderUsageDescription':
            'MathToWord сохраняет готовые документы .docx на Рабочий стол.',
        'NSDocumentsFolderUsageDescription':
            'MathToWord может сохранять документы .docx.',
    },
)
