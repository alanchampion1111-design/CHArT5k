@echo off
taskkill /f /im ngrok.exe >nul 2>&1
ngrok start --config "C:\CHArT5k-Puppet\ngrok.yml" puppet-bridge
pause