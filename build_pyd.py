from distutils.core import setup
from Cython.Build import cythonize

setup(
    name='Anything you want',
    ext_modules=cythonize(["common/excel/Array.py","common/excel/Report.py","common/excel/Template.py","common/excel/Write.py", 
                           "common/http/Format.py", "common/http/Http.py", 
                           "common/init/Init.py", "common/init/InitColumn.py","common/init/InitConfig.py","common/init/InitFile.py","common/init/InitSheet.py",
                           "common/ui/ComboCheckBox.py", "common/ui/TextEdit.py", "common/ui/Ui_mainWindow.py", 
                           "common/utils/analy.py","common/utils/Log.py","common/utils/Util.py"
                           ], language_level=3
        ),
)