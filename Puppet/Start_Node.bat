@echo off
cd C:\CHArT5k-Puppet
echo Starting CHArT5k-Puppet via PM2...
npm start
echo App started in background!
echo To monitor node...
npm run logs
pause
