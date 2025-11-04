#打包程序
pyinstaller -F -w -i logo.ico --hidden-import=pandas --hidden-import=openpyxl main_interface.py  
