@if not exist "%0\..\bin\Release\OutlookRuleMgr.exe" call "%0\..\build.bat"
@"%0\..\bin\Release\OutlookRuleMgr.exe" %*
