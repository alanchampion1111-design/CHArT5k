@echo off
title Node
cd C:\CHArT5k-Puppet
echo Starting CHArT5k-Puppet via PM2...
npm start & npm run logs
pause
