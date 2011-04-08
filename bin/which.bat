@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

SET OPT_ALL=0
SET OPT_HELP=0

:PARSEOPT
IF "%1" == "-h" (
  SET OPT_HELP=1
  SHIFT
  GOTO :PARSEOPT
)
IF "%1" == "--help" (
  SET OPT_HELP=1
  SHIFT
  GOTO :PARSEOPT
)
IF "%1" == "-a" (
  SET OPT_ALL=1
  SHIFT
  GOTO :PARSEOPT
)

IF "%1" == "" (
  echo usage: which [options] filename
  echo   -h --help    print this text
  echo   -a           print all matching pathname of each argument
  GOTO :EOF
)

SET NAME=%1

SET DOTPATH=.;%PATH%

FOR %%i IN ("%DOTPATH:;=" "%") DO (
  SET DIR=%%~i
  IF EXIST "!DIR!\!NAME!" (
    echo !DIR!\!NAME!
    IF NOT "!OPT_ALL!" == "1" (
      GOTO :EOF
    )
  )
  FOR %%j IN ("%PATHEXT:;=" "%") DO (
    SET EXT=%%~j
    IF EXIST "!DIR!\!NAME!!EXT!" (
      echo !DIR!\!NAME!!EXT!
      IF NOT "!OPT_ALL!" == "1" (
        GOTO :EOF
      )
    )
  )
)

