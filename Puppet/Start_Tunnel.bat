@echo off
title CHArT5k-Puppet Tunnel Manager
echo Cleaning up existing ngrok instances...
taskkill /f /im ngrok.exe >nul 2>&1
cd C:\CHArT5k-Puppet
echo Starting ngrok tunnel using config: ngrok-harmless.yml
echo ...from GDrive:\IT Business\Parkrun\CHArT5k-Puppet\ (default on a single IP address)
ngrok start --all --config ngrok-harmless.yml
echo Alternative "local" club assignments (who?) where a different yml may operate:
echo Poole AC Parkrunners (Keren):             ngrok start --all --config ngrok-doormat.yml 
echo Witney Road Runners (Joan):               ngrok start --all --config ngrok-harmless.yml
echo Christchurch Runners (Micah):             ngrok start --all --config ngrok-shivering.yml
echo Overton Harriers & AC Clubrunners (Dave): tbd
pause
