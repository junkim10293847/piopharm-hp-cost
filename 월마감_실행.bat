@echo off
chcp 65001 > nul
REM 사용법: data\inbox 폴더에 이번 달 기초 파일을 넣고 이 파일을 더블클릭합니다.
REM 결과는 data\inbox\output 폴더에 저장됩니다.
cd /d %~dp0
if not exist data\inbox mkdir data\inbox
node cli\hp-cost.js data\inbox
pause
