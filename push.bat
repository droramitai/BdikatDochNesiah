@echo off
cd /d C:\Projects\Ituran
echo.
echo ===== דוחף לגיטהאב =====
git add -A
git status
echo.
set /p MSG="הכנס הודעת קומיט (או Enter לברירת מחדל): "
if "%MSG%"=="" set MSG=update
git commit -m "%MSG%"
git push origin main
echo.
echo ===== הועלה בהצלחה! =====
echo Streamlit Cloud יתעדכן תוך 1-2 דקות
echo.
pause
