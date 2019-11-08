@echo off

SET scriptDir=%~dp0

set setupFile=%scriptDir%setup.ps1

powershell.exe -noexit -file %setupFile%
