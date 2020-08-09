@echo off
setlocal enabledelayedexpansion

echo **************************
echo *     ファイル名置換     *
echo **************************
set /p beforeFileName="変更部のファイル名を入力してください："
set /p afterFileName="変更後のファイル名を入力してください："

if ""%beforeFileName%"" == """" (
  echo 変更部のファイル名入力漏れ
  goto Jump
)

for %%f in ( * ) do (
  set newFileName=%%f
  set newFileName="!newFileName:%beforeFileName%=%afterFileName%!"
  ren "%%f" !newFileName!
)

:Jump
pause

endlocal 
