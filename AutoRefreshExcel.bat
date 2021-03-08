if "%1"=="hide" goto CmdBegin
start mshta vbscript:createobject("wscript.shell").run("""%~0"" hide",0)(window.close)&&exit
:CmdBegin
@echo # Write by zhengwee on 2021.3.4  >refresh.pyw


@echo # refresh all excel connection and save. >>refresh.pyw
@echo import win32com.client  >>refresh.pyw
@echo excelpath = r"%~dp0\CD_excess_order_daily_To.xlsx" >>refresh.pyw
@echo xlapp = win32com.client.DispatchEx("Excel.Application") >>refresh.pyw
@echo wb = xlapp.Workbooks.Open(excelpath) >>refresh.pyw
@echo wb.RefreshAll() >>refresh.pyw
@echo xlapp.CalculateUntilAsyncQueriesDone() >>refresh.pyw
@echo wb.Save() >>refresh.pyw
@echo xlapp.Quit() >>refresh.pyw
@echo off  
C:  
cd cd /d %~dp0
start pythonw refresh.pyw
@ping 127.0.0.1 -n 10 >nul
del /a /f /s refresh.pyw
exit  
