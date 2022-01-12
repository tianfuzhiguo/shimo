from distutils.core import setup
from Cython.Build import cythonize

setup(
    name='Anything you want',
    ext_modules=cythonize(["common/excel/Array.py","common/excel/Report.py","common/excel/Template.py","common/excel/Write.py", 
                           "common/http/Format.py", "common/http/Http.py", 
                           "common/init/Init.py","common/init/InitConfig.py","common/init/InitExcel.py",
                           "common/ui/ComboCheckBox.py", "common/ui/TextEdit.py", "common/ui/Ui_mainWindow.py", 
                           "common/utils/Analy.py","common/utils/SmLog.py","common/utils/Util.py","common/utils/ExcelUtil.py",
                           ], language_level=3
        ),
)