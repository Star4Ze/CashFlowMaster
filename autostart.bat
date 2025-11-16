@echo off
echo Запускаем Python-бота...
cd /d "C:\Users\HomePC\CashFlowMaster"
python -m pip install -r requirements.txt --quiet
python bot.py
pause