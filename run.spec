# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['run.py'],
             pathex=['D:\\TRPG\\TRPG_replay_video_generator', 'D:\\TRPG\\TRPG_replay_video_generator\\GUI', 'D:\\TRPG\\TRPG_replay_video_generator\\Core'],
             binaries=[],
             datas=[],
             hiddenimports=['PyQt5.sip', 'pptx.sip', 'six.sip'],
             hookspath=['D:\\TRPG\\backup'],
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
          name='pptx-generator',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          icon = 'D:\TRPG\TRPG_replay_video_generator\icon.ico',
          console=False )
