@pushd %~dp0
@git.exe checkout -b release
@git.exe checkout release
@git.exe merge master --no-commit
@rd /s /q bin
@rd /s /q obj
@rd /s /q packages
@call build.bat
@git.exe add --force bin
@git.exe commit -m "Build release"
@popd
