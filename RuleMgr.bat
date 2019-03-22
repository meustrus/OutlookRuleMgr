@if not exist "%~dp0\bin\Release\OutlookRuleMgr.exe" call "%~dp0\build.bat"
@"%~dp0\bin\Release\OutlookRuleMgr.exe" %*
