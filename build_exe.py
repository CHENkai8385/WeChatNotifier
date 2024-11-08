import PyInstaller.__main__

PyInstaller.__main__.run([
    'wxauto.py',
    '--name=WeChatNotifier',
    '--onefile',
    '--noconsole',
    '--icon=assets/icon.ico',
    '--add-data=assets/images/wechat.png;assets/images',
    '--hidden-import=pystray._win32',
])