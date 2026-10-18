# TB_SCRIPT

Python and JavaScript scripts that build the daily and monthly TB programme reports for Mahisagar district.

## What it does

- Reads the NIKSHAY/Spectrum export files
- Builds the notification, TPT, PTFU and DTC reports
- Builds the district and monthly PowerPoint decks
- Runs daily at 10 AM through `run_daily.bat` and `daily_master.py`

## Notes

- Data files (Excel, CSV, PowerPoint, PDF) and logs are not stored here. They are excluded by `.gitignore`.
- See `SETUP_ONCE.txt` for first-time setup on a new PC.