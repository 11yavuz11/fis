from PyQt5 import uic

with open('pencere.ui', 'w', encoding="utf-8") as fout:
   uic.compileUi('pencere.py', fout)