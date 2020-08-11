@echo

cd C:\Users\xxx\workspace\1.3.0.20200723

python build_pyd.py build_ext --inplace

pyinstaller -F -w -i 1.ico sm.spec
pause