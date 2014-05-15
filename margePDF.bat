
cd /d %~dp0

set CURRENT_DIR=%~dp0

set CONCATPDF_DIR=ConcatPDF
set CONCATPDF_BAT=%CONCATPDF_DIR%\ConcatPDF.bat

del %CONCATPDF_DIR%\*.pdf

IF "%1" == "" (
  SET CONCATPDF_OUT=outfile.pdf
) ELSE (
  SET CONCATPDF_OUT=%1
)

for %%i in (*.pdf) do (
  echo %%i
  move /Y %%i %CONCATPDF_DIR%
)

call %CONCATPDF_BAT% %CONCATPDF_OUT%

move /Y %CONCATPDF_OUT% %~dp0\.
