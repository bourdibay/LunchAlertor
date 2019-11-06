@echo off

REM This file is called by our Windows Task Scheduler.
REM This script is executed at workstation unlock. This script lets us run the python app only when we unlock the computer between 12h and 13h.

setlocal enableextensions enabledelayedexpansion

set currentHour=%time:~0,2%
set script=%1
set startHour=%2
set endHour=%3

if "%currentHour%" geq "%startHour%" if "%currentHour%" leq "%endHour%" goto run

goto end

:run
  "D:\Program Files\Python\Python37\python.exe" "%script%" "%startHour%" "%endHour%"

:end
