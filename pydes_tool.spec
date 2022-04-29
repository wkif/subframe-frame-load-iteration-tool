# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['pydes_tool.py'],
             pathex=['F:\\File_my\\Project\\Plug-in_1\\Pyqt',
             'F:\\File_my\\Project\\Plug-in_1\\Pyqt\\venv\\Lib\\site-packages\\PyQt5',
             'F:\\File_my\\Project\\Plug-in_1\\Pyqt\\venv\\Lib\\site-packages\\PyQt5\\Qt5\\bin',
             'F:\\File_my\\Project\\Plug-in_1\\Pyqt\\venv\\Lib\\site-packages\\PyQt5\\Qt5\\plugins',
             'C:\\Windows\\System32\\downlevel'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
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
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='pydes_tool',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='pydes_tool.ico')
