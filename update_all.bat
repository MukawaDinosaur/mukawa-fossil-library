@echo off

chcp 65001 > nul

title 鵡川高校 化石デジタルライブラリー 自動更新

echo.
echo ============================================================
echo  鵡川高校 化石デジタルライブラリー
echo  自動更新
echo ============================================================
echo.
echo Excel・画像のチェック
echo JSON生成
echo サムネイル生成
echo Git commit
echo GitHub push
echo を開始します。
echo.

cd /d "%~dp0"

python update_all.py

echo.
echo ============================================================
echo  処理が終了しました。
echo ============================================================
echo.

pause
