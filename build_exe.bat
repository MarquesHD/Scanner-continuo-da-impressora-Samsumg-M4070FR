@echo off
echo ============================================
echo  Instalando dependencias...
echo ============================================
pip install -r requirements.txt
pip install pyinstaller

echo.
echo ============================================
echo  Gerando o executavel (.exe)...
echo ============================================
pyinstaller --noconfirm --onefile --windowed ^
  --name "ScannerSamsung" ^
  --hidden-import=win32com ^
  --hidden-import=win32com.client ^
  main.py

echo.
echo ============================================
echo  Pronto! O executavel esta em:
echo  dist\ScannerSamsung.exe
echo ============================================
pause
