@echo off

echo **************************
echo *     ファイル名追加     *
echo **************************
echo 0:先頭
echo 1:末尾

set /p select="数字で選択して下さい："
set /p add="追記する文字を入力して下さい："

if ""%add%"" == """" (
  echo 追記文字入力漏れ
  goto Jump
)

if %select%==0 (
    for %%i in (*) do (
        rem 実行バッチファイル自身以外に実行
        if not %%i==%~n0%~x0 (
            rem 先頭に追記
            ren "%%i" "%add%%%~ni%%~xi"
        )
    )
) else (
    for %%i in (*) do (
        rem 実行バッチファイル自身以外に実行
        if not %%i==%~n0%~x0 (
            rem 末尾に追記
            ren "%%i" "%%~ni%add%%%~xi"
        )
    )
)

:Jump
pause

endlocal 