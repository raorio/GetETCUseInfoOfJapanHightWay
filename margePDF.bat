
cd /d %~dp0

set CURRENT_DIR=%~dp0

set CONCATPDF_DIR=ConcatPDF
set CONCATPDF_BAT=%CONCATPDF_DIR%\ConcatPDF.bat
set CONCATPDF_OUT=%CONCATPDF_DIR%\outfile.pdf

del %CONCATPDF_DIR%\*.pdf

for %%i in (*.pdf) do (
  echo %%i
  move /Y %%i %CONCATPDF_DIR%
)

call %CONCATPDF_BAT% %CONCATPDF_OUT%

move /Y %CONCATPDF_OUT% .
