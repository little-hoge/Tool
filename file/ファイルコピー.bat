@echo off
setlocal enabledelayedexpansion

echo **************************
echo *     ファイルコピー     *
echo **************************
set /p filename="ファイル名を入力してください："
set /p number="ファイル数を入力してください："
set /p extension="拡張子を入力してください："


if ""%filename%"" == """" (
  echo ファイル名入力漏れ
  goto Jump
)

if ""%number%"" == """" (
  echo ファイル数入力漏れ
  goto Jump
)

if not ""%extension%"" == """" (
  set extension=.%extension%
)

for /l %%n in (1,1,%number%) do (
  set num=%%n
  set copy_target=%filename%!extension!
  set copy_source=%filename%_!num!!extension!

rem 確認コマンド入力
  if not ""%extension%"" == """" (
  echo F |xcopy !copy_target! !copy_source!

) else (
  echo D |xcopy !copy_target! !copy_source!

)
  
)
 
:Jump

endlocal 
pause
