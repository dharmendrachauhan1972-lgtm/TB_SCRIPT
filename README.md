# TB_SCRIPT

Python and JavaScript scripts that produce the daily and monthly TB reports for Mahisagar district.

## What it does

The scripts read the NIKSHAY/Spectrum exports and build the notification, TPT, PTFU and DTC reports, plus the district and monthly PowerPoint decks. `run_daily.bat` starts `daily_master.py` every day at 10 AM.

## Notes

Excel, CSV, PowerPoint and PDF files and the logs are not stored in this repo, because `.gitignore` excludes them. For first-time setup on a new PC, see `SETUP_ONCE.txt`.