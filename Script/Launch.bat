@echo off
SET ScriptDir="%~dp0
SET PSScriptPath=%ScriptDir%SSWLoginScript.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File %PSScriptPath%
