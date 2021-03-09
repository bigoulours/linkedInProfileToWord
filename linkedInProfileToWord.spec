# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['linkedInProfileToWord.py'],
             pathex=['.'],
             binaries=[],
             datas=[('images/linkedin.ico', 'images/'), ('resources/profile_template_DE_jinja2.docx', 'resources/'),
			('resources/profile_template_EN_jinja2.docx', 'resources/'), ('resources/profile_template_FR_jinja2.docx', 'resources/')],
             hiddenimports=[],
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
          name='linkedInProfileToWord',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , icon='images/linkedin.ico')
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='linkedInProfileToWord')
