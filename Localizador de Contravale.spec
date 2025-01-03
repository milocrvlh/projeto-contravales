# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['Localizador de Contravale.py'],
             pathex=['C:\\Users\\milo\\Desktop\\Projeto Contravales'],
             binaries=[],
             datas=[],
             hiddenimports=['numpy', 'numpy.core', 'numpy.lib', 'openpyxl', 'python-dateutil', 'pytz', 'six', 'et-xmlfile', 'jdcal'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='Localizador de Contravale',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='Localizador de Contravale')
