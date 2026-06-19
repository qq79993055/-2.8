# -*- mode: python ; coding: utf-8 -*-
# Mac版打包配置 — 生成 .app（GUI应用）
# 运行：pyinstaller mac_build.spec

a = Analysis(
    ['photo_report_gui手机版.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('单页模版.pptx', '.'),
        ('装饰底图.jpg', '.'),
        ('issues_db.json', '.'),
        ('remedies_db.json', '.'),
        ('settings.json', '.'),
    ],
    hiddenimports=['pptx', 'PIL', 'PIL._imaging', 'PIL._imagingtk'],
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
    name='华为门店验收工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,        # Mac：False = 不弹终端窗口
    windowed=True,        # Mac：生成 GUI 应用
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,      # 自动选择当前架构（Intel 或 Apple Silicon）
    codesign_identity=None,
    entitlements_file=None,
)

# Mac 上用 COLLECT + BUNDLE 生成 .app
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name='华为门店验收工具',
)

app = BUNDLE(
    coll,
    name='华为门店验收工具.app',
    icon=None,
    bundle_identifier='com.yongge.huawei.inspection',
)
