@echo off
:: --- 設定（必要に応じて変更） ---

:: 1. Anacondaインストール先（環境に応じて書き換え）
set ANACONDA_PATH=C:\Users\chopi\anaconda3

:: 2. 実行する仮想環境名（通常は base）
set ENV_NAME=base

:: 3. 実行するスクリプトのパス（相対）
set SCRIPT_PATH=script\invoice_generator.py

:: --- 実行処理 ---

:: Anaconda環境をアクティベート
call "%ANACONDA_PATH%\Scripts\activate.bat" %ENV_NAME%

:: スクリプト実行
python "%SCRIPT_PATH%"

:: 実行終了確認
echo.
pause
