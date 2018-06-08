@echo off
SET ScriptDir="%~dp0
SET PSScriptPath=%ScriptDir%SSWTemplateScript.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File %PSScriptPath%
