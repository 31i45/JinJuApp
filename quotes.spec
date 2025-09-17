# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 分析应用程序依赖
analysis = Analysis(
    ['quotes.py'],
    pathex=[],
    binaries=[],
    datas=[('quotes.json', '.'), ('logo.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'sqlite3', 'cryptography', 'psutil',  # 排除不需要的模块
             'pandas', 'numpy', 'matplotlib', 'scipy',
             'PIL', 'urllib3', 'requests'],
    # 移除了不再支持的参数 win_no_prefer_redirects 和 win_private_assemblies
    cipher=block_cipher,
    noarchive=False,
)

# 创建PYZ文件
pyz = PYZ(analysis.pure, analysis.zipped_data, cipher=block_cipher)

# 配置可执行文件
exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.zipfiles,
    analysis.datas,
    [],
    name='人民日报金句',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,  # 禁用符号表剥离，Windows系统默认没有
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo.ico',
)