@echo off
setlocal enabledelayedexpansion

echo **************************
echo *     �t�@�C�����u��     *
echo **************************
set /p beforeFileName="�ύX���̃t�@�C��������͂��Ă��������F"
set /p afterFileName="�ύX��̃t�@�C��������͂��Ă��������F"

if ""%beforeFileName%"" == """" (
  echo �ύX���̃t�@�C�������͘R��
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
