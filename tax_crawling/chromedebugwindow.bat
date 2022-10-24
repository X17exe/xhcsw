@echo off
if "%1" == "h" goto begin
mshta vbscript:createobject("wscript.shell").run("%~nx0 h",0)(window.close)&&exit
:begin
C:
cd C:\Program Files\Google\Chrome\Application
chrome.exe --remote-debugging-port=9000 --user-data-dir="C:\chromedebug"
