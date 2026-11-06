# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[
        ('D:\\Program Files\\GDEClaw\\Library\\bin\\ffi.dll', '.'),
        ('D:\\Program Files\\GDEClaw\\Library\\bin\\libexpat.dll', '.'),
        ('D:\\Program Files\\GDEClaw\\Library\\bin\\libcrypto-3-x64.dll', '.'),
        ('D:\\Program Files\\GDEClaw\\Library\\bin\\libssl-3-x64.dll', '.'),
        ('D:\\Program Files\\GDEClaw\\Library\\bin\\liblzma.dll', '.'),
        ('D:\\Program Files\\GDEClaw\\Library\\bin\\libbz2.dll', '.'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\lib\\Microsoft.Web.WebView2.Core.dll', 'webview/lib'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\lib\\Microsoft.Web.WebView2.WinForms.dll', 'webview/lib'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\lib\\WebBrowserInterop.x64.dll', 'webview/lib'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\lib\\WebBrowserInterop.x86.dll', 'webview/lib'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\lib\\runtimes\\win-x64\\native\\WebView2Loader.dll', 'webview/lib/runtimes/win-x64/native'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\lib\\runtimes\\win-x86\\native\\WebView2Loader.dll', 'webview/lib/runtimes/win-x86/native'),
        # Playwright 驱动：node.exe（Node 运行时）+ package（Playwright JS 核心），
        # 供「平台网页配置 -> 一键登录」启动浏览器抓取 Cookie 使用。
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\playwright\\driver\\node.exe', 'playwright/driver'),
    ],
    datas=[
        ('index.html', '.'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\js\\api.js', 'webview/js'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\js\\customize.js', 'webview/js'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\js\\finish.js', 'webview/js'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\js\\lib\\dom_json.js', 'webview/js/lib'),
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\webview\\js\\lib\\polyfill.js', 'webview/js/lib'),
        # Playwright 驱动 JS 核心（cli.js + lib + node_modules 等）
        ('D:\\Program Files\\GDEClaw\\lib\\site-packages\\playwright\\driver\\package', 'playwright/driver/package'),
    ],
    hiddenimports=[
        'playwright',
        'greenlet',
        # sync API（一键登录抓 cookie 使用）
        'playwright.sync_api',
        'playwright.sync_api._generated',
        'playwright.sync_api._context_manager',
        'playwright._impl._greenlets',
        'playwright._impl._sync_base',
        'playwright._impl._async_base',
        'playwright._impl._assertions',
        # async API（保留兼容）
        'playwright.async_api',
        'playwright.async_api._generated',
        'playwright.async_api._context_manager',
        'playwright._impl._driver',
        'playwright._impl._transport',
        'playwright._impl._connection',
        'playwright._impl._object_factory',
        'playwright._impl._api_structures',
        'playwright._impl._helper',
        'playwright._impl._errors',
        'playwright._impl._impl_to_api_mapping',
        'playwright._repo_version',
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
    [],
    exclude_binaries=True,
    name='MRDN系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MRDN系统',
)
