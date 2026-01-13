@echo off
CALL conda activate pdfconverter
cd /d %~dp0
python pdf_converter_gui.py
pause

