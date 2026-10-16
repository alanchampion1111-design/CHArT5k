@echo off
start "" "C:/CHArT5k-Puppet/Run_Tunnel.bat"
start "" "C:/CHArT5k-Puppet/Run_Node.bat"
timeout /t 3
powershell.exe -ExecutionPolicy Bypass -File "C:/CHArT5k-Puppet/MoveToDesktop2.ps1"