@echo off
setlocal enabledelayedexpansion

:: 現在のディレクトリ名を取得
for %%f in (%cd%) do set foldername=%%~nxf

:: 現在の日付をYYYYMMDD形式で取得（WMICの代替としてPowerShellを使用）
for /f %%a in ('powershell -Command "Get-Date -Format yyyyMMdd"') do set today=%%a

:: 初期ファイル名を設定
set "filename=%foldername%_%today%.txt"

:: ファイルの重複回避
set count=1
:checkfile
if exist "%filename%" (
    set /a count+=1
    set "filename=%foldername%_%today%_(%count%).txt"
    goto checkfile
)

:: ファイルを作成して内容を書き込む
echo %foldername% %today% > "%filename%"

:: 作成したファイルを開く（使用するエディタを変更可能）
start notepad "%filename%"

endlocal
