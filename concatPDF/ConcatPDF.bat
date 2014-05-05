
cd /d %~dp0

REM SET TARGET_VERSION=1.2.5-windows732bit
SET TARGET_VERSION=1.2.5-windows764bit
SET CONCATPDF_EXE=%TARGET_VERSION%\ConcatPDF.exe
IF %1 == "" (
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

%CONCATPDF_EXE% /outfile %OUT_FILE% %IN_FILES%

endlocal
