@echo off
echo Installing required packages...
pip install -r requirements.txt

echo.
echo Starting PDF to Word Converter...
echo Server will run at: http://localhost:5000
echo.
echo Make sure you have these tools in the 'tools' folder:
echo - poppler/bin/
echo - tesseract/
echo - java/bin/
echo - tabula.jar
echo.

python server.py
pause
