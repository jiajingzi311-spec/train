@echo off
chcp 65001 >nul
title MRDN系统 一键打包
cd /d "%~dp0"
echo.
echo 正在打包（会自动备份和恢复你的已配置数据）...
echo.
python 打包.py
echo.
pause
