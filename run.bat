@echo off
title DTDC Bulk Parcel Tracker
echo ========================================================
echo            DTDC BULK PARCEL TRACKER DASHBOARD
echo ========================================================
echo.
echo Starting server...
start "" "http://127.0.0.1:5000"
python app.py
pause
