# -*- mode: python -*-
a = Analysis(['cemtimer.py'],
             pathex=['C:\\Users\\Chris\\Google Drive\\PycharmProjects\\cemtimer'],
             hiddenimports=[],
             hookspath=None,
             runtime_hooks=None)
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='cemtimer.exe',
          debug=False,
          strip=None,
          upx=True,
          console=False , icon='clock.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=None,
               upx=True,
               name='cemtimer')
