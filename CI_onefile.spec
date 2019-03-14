# -*- mode: python -*-

block_cipher = None

from pathlib import Path

a = Analysis([Path(SPEC).parent.joinpath('Computer Info.py')],
             pathex=[Path(SPEC).parent.joinpath('Computer Info.py')],
             binaries=[],
             datas=[],
             hiddenimports=['win10toast'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

a.datas += [('logo.ico', str(Path(SPEC).parent.joinpath('logo.ico')), 'Data')]
a.datas += [('multi_comp_settings.cfg', str(Path(SPEC).parent.joinpath('multi_comp_settings.cfg')), 'Data')]
a.datas += [('other_applications.prg', str(Path(SPEC).parent.joinpath('other_applications.prg')), 'Data')]
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Computer Info',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True,
          icon=str(Path(SPEC).parent.joinpath('logo.ico')))
