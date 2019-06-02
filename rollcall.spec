# -*- mode: python -*-

block_cipher = None


a = Analysis(['rollcall.py'],
             pathex=['C:\\Users\\Administrator\\Desktop\\µãÃûÈí¼þ'],
             binaries=[],
             datas=[],
             hiddenimports=['pkg_resources'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='rollcall',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='image\\icon.ico')
