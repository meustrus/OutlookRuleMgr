@pushd "%0\.."
NuGet.exe restore OutlookRuleMgr.sln
"%PROGRAMFILES(x86)%\MSBuild\14.0\Bin\msbuild.exe" OutlookRuleMgr.sln /t:Clean,Rebuild /p:Configuration=Release
@popd
