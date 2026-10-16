@echo off
:: 🚀 Start Node (with log) on the LHS (60%), with specific Tunnel on the RHS (40%)
wt.exe --maximized ^
  cmd.exe /k "C:\CHArT5k-Puppet\Start_Node.bat" ; ^
  split-pane -v -s 0.4 cmd.exe /k "C:\CHArT5k-Puppet\Start_Tunnel.bat"

:: Give the maximized container 4 seconds to register its window handle
timeout /t 4

:: 💻 Teleport the combined window container to Desktop 2
powershell.exe -ExecutionPolicy Bypass -Command " ^
  Import-Module VirtualDesktop; ^
  Get-Process -Name 'WindowsTerminal' | ^
    Where-Object { \$_.MainWindowHandle -ne 0 } | ^
    ForEach-Object { ^
      Move-Window ` ^
        -Hwnd $_.MainWindowHandle ` ^
        -Desktop (Get-Desktop -Index 1) ^
    } ^
"
echo Dashboard deployed to Desktop 2!
