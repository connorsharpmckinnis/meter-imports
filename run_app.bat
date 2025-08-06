@echo off
cd /d %~dp0
call .venv\Scripts\activate
start http://localhost:8866
voila Meter_Import.ipynb --no-browser --strip_sources=True --Voila.file_whitelist=".*"