@echo off
setlocal enabledelayedexpansion

echo **************************
echo *     �t�@�C���R�s�[     *
echo **************************
set /p filename="�t�@�C��������͂��Ă��������F"
set /p number="�t�@�C��������͂��Ă��������F"
set /p extension="�g���q����͂��Ă��������F"


if ""%filename%"" == """" (
  echo �t�@�C�������͘R��
  goto Jump
)

if ""%number%"" == """" (
  echo �t�@�C�������͘R��
  goto Jump
)

if not ""%extension%"" == """" (
  set extension=.%extension%
)

for /l %%n in (1,1,%number%) do (
  set num=%%n
  set copy_target=%filename%!extension!
  set copy_source=%filename%_!num!!extension!

rem �m�F�R�}���h����
  if not ""%extension%"" == """" (
  echo F |xcopy !copy_target! !copy_source!

) else (
  echo D |xcopy !copy_target! !copy_source!

)
  
)
 
:Jump

endlocal 
pause
