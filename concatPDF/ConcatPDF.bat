
cd /d %~dp0

SET CONCATPDF_64BIT=C:\Program Files\ConcatPDF
SET CONCATPDF_32BIT=C:\Program Files (x86)\ConcatPDF

SET PATH=%PATH%;%CONCATPDF_64BIT%;%CONCATPDF_32BIT%
SET CONCATPDF_EXE=ConcatPDF.exe
IF "%1" == "" (
  SET OUT_FILE=outfile.pdf
) ELSE (
  SET OUT_FILE=%1
)


SET IN_FILES=
SET PREV_IN_FILE=

setlocal enabledelayedexpansion
FOR %%i IN (*.pdf) DO (
  IF NOT %OUT_FILE% == %%i SET IN_FILES=%%i !IN_FILES!
)

echo %PATH%
%CONCATPDF_EXE% /outfile %OUT_FILE% %IN_FILES%

endlocal
