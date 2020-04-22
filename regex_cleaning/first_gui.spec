# -*- mode: python ; coding: utf-8 -*
block_cipher = None

import sys
sys.setrecursionlimit(5000)

def get_pandas_path():
    import pandas
    pandas_path = pandas.__path__[0]
    return pandas_path

a = Analysis(['first_gui.py'],
             pathex=['C:\\Users\\Cat Mai\\Documents\\Work\\NGSX\\Coding\\NGSX\\regex_cleaning'],
             binaries=[],
             datas=[],
             hiddenimports=["pkg_resources.py2_warn"],
             hookspath=None,
             runtime_hooks=None,
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

dict_tree = Tree(get_pandas_path(), prefix='pandas', excludes=["*.pyc"])
a.datas += dict_tree
a.binaries = filter(lambda x: 'pandas' not in x[0], a.binaries)

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='first_gui',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False )
scoll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=False,
               upx_exclude=[],
               name='first_gui')
