@echo off

echo **************************
echo *     �t�@�C�����ǉ�     *
echo **************************
echo 0:�擪
echo 1:����

set /p select="�����őI�����ĉ������F"
set /p add="�ǋL���镶������͂��ĉ������F"

if ""%add%"" == """" (
  echo �ǋL�������͘R��
  goto Jump
)

if %select%==0 (
    for %%i in (*) do (
        rem ���s�o�b�`�t�@�C�����g�ȊO�Ɏ��s
        if not %%i==%~n0%~x0 (
            rem �擪�ɒǋL
            ren "%%i" "%add%%%~ni%%~xi"
        )
    )
) else (
    for %%i in (*) do (
        rem ���s�o�b�`�t�@�C�����g�ȊO�Ɏ��s
        if not %%i==%~n0%~x0 (
            rem �����ɒǋL
            ren "%%i" "%%~ni%add%%%~xi"
        )
    )
)

:Jump
pause

endlocal 