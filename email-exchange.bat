@echo off
cd /d "%~dp0"
pipenv run python email-exchange.py
