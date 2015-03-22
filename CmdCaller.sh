@echo off
goto :cmd
:cmd
call cmd.exe
goto :looper
:looper
goto :cmd
