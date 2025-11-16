Set WShell = CreateObject("WScript.Shell")
WShell.CurrentDirectory = "C:\Users\HomePC\CashFlowMaster"
WShell.Run "cmd.exe /c python -m pip install -r requirements.txt --quiet && python bot.py && pause", 2, False
