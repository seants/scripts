# ical_to_xlsx.py
This script roughly mimics what the website https://www.gcal2excel.com/ does for a Google Calendar: takes a specific time range and exports all events within that range in that calendar as a .xlsx file with date, name, and duration. Since it uses the ics Calendar package, it is easy to modify the output fields to include/exclude the standard Event fields (description/start/end).

Originally, there was a column that did cost-calculations based on the event duration and description, but that's been omitted now for generality.

To run, download the calendar ICS file from Google (it will come packaged in a zip file), modify the `ZIP_DIRECTORY` field with the location of the download, and run the script from the destination directory (where you want the xlsx file to be saved). By default, all events from the previous month are listed, but the month can be adjusted with command line parameters in the form of `python ical_to_xlsx.py 2016 7` for example.

To install the dependencies, run `pip install ics openpyxl`.
