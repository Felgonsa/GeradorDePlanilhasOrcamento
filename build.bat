@echo off
echo Apagando versao antiga...
rmdir /s /q dist
rmdir /s /q build
del gerar_excel.spec

echo.
echo Gerando novo executavel...
pyinstaller --onefile --noconsole --icon=icone.ico gerar_excel.py

echo.
echo Build finalizado!
pause
