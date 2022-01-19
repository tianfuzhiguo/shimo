# -*- mode: python ; coding: utf-8 -*-

import sys
sys.setrecursionlimit(5000)

block_cipher = None


a = Analysis(['jenkins.py'],
             pathex=['C:\\Users\\sudi\\workspace\\20201106'],
             binaries=[],
             datas=[('source/template','source'),('source/2.png','source'),('source/1.ico','source')],
             hiddenimports=['common.excel.Array','common.excel.Report','common.excel.Template','common.excel.Write', 
                            'common.http.Format', 'common.http.Http', 
                            'common.init.Init','common.init.InitConfig','common.init.InitExcel',
                            'common.ui.ComboCheckBox', 'common.ui.TextEdit', 'common.ui.Ui_mainWindow',
                            'common.utils.Analy','common.utils.SmLog','common.utils.ExcelUtil','common.utils.Util',
                            'xlutils','xlutils.copy','xlwt','xlrd','xlsxwriter','demjson','xmltodict','configparser','shutil','PyQt5.QtWidgets','PyQt5',
                            'cx_Oracle','pymysql','pymssql',
                            'apscheduler.schedulers.background','apscheduler.triggers.date','PyQt5.QtCore','yagmail',
                            'openpyxl','jsonschema',
                            'requests'],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='jenkins',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='source/1.ico')
