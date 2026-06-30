# lab equipment calibration tracker

built this after my co-op at ORTECH because we were tracking all our equipment calibration on paper and someone had to manually type it into excel. seemed like a waste of time and also easy to miss stuff

at ORTECH i was collecting samples every 30 minutes during leaching runs and submitting 600+ samples to the ICP lab. all of that was tracked by hand. the pH meter and flow meter both needed regular calibration and i honestly dont know how we kept track of it without something like this

## what it does

- reads equipment data from a csv (equipment.csv)
- calculates when each instrument is next due for calibration
- flags anything thats overdue or coming up in the next 2 weeks
- prints a summary to the terminal
- exports a colour coded excel report

## how to run it

make sure you have openpyxl installed first

```
pip install openpyxl
```

then just run

```
python calibration_tracker.py
```

it'll print the status and save calibration_report.xlsx in the same folder

## the csv format

equipment_id, name, location, last_calibrated (YYYY-MM-DD), calibration_interval_days, responsible, notes

you can just edit the csv to add or remove equipment

## things i want to add eventually

- email alerts when something goes overdue
- log when calibrations actually get done (right now you have to manually update the csv)
- maybe a simple UI so the lab manager doesnt have to touch the csv

