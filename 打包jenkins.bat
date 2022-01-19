@echo

cd C:\Users\sudi\workspace\20201106

python build_pyd.py build_ext --inplace

pyinstaller -F -w -i 1.ico jenkins.spec

pause