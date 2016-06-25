import arrow
import csv
from dateutil import tz
from ics import Calendar
import openpyxl
import os
import os.path
import sys
import zipfile

# TODO: fail immediately if filename already exists

ZIP_DIRECTORY = 'E:\\Downloads'
ZIP_ICAL_STR = '.ical.zip'
OUT_FILE_TYPE = '.xlsx'

SECONDS_PER_MINUTE = 60
MINUTES_PER_HOUR = 60

COST_PER_HOUR = 1

def get_desired_date():
    if len(sys.argv) == 1:
        return arrow.now().replace(months=-1)
    elif len(sys.argv) == 3:
        return arrow.get(int(sys.argv[1]), int(sys.argv[2]), 1, tzinfo=tz.tzlocal())
    else:
        raise Exception('Unrecognized number of command-line arguments: %i; expected 0 or 2' % (len(sys.argv) - 1))

def get_zip_file_name():
    files = os.listdir(ZIP_DIRECTORY)
    candidate_files = [filename for filename in files if ZIP_ICAL_STR in filename]
    if len(candidate_files) == 1:
        zip_file_name = candidate_files[0]
    elif len(candidate_files) == 0:
        raise Exception('No .ical ZIP file found. Try downloading again, and make sure you are in the right directory: %s.' % ZIP_DIRECTORY)
    else:
        raise Exception('Multiple .ical ZIP files found in %s. Please delete all old .ical ZIP files.' % ZIP_DIRECTORY)
    return os.path.join(ZIP_DIRECTORY, zip_file_name)

def get_ical_file_name(zip_file):
    ical_file_names = zip_file.namelist()
    if len(ical_file_names) != 1:
        raise Exception('ZIP archive had %i files; expected 1.' % len(ical_file_names))
    return ical_file_names[0]

def filter_by_date(events, date):
    return events[date.floor('month'):date.ceil('month')]

def filter_by_custom(events):
    return [event for event in events if event.name.lower() != 'eva out']
    
def strip_number(events):
    for event in events:
        event.name = ''.join(filter(lambda a: not a.isnumeric(), event.name)).strip()

def calculate_cost(event):
    return event.duration.seconds / SECONDS_PER_MINUTE / MINUTES_PER_HOUR * COST_PER_HOUR

def write_output(events, date):
    name = date.format('YYYY MMMM') + OUT_FILE_TYPE
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    for i, event in enumerate(events):
        worksheet.append([event.name, event.begin.format('MMMM DD'), event.duration, calculate_cost(event)])
        worksheet["D%i" % (i + 1)].number_format
    worksheet.append(["", "", "", "=SUM(D1:D%i)" % len(events)])
    workbook.save(name)

def get_desired_events(ical_file, date):
    calendar = Calendar(ical_file.read().decode('iso-8859-1'))
    events = filter_by_date(calendar.events, date)
    events = filter_by_custom(events)
    strip_number(events)
    return events

def main():
    desired_date = get_desired_date()
    # We need two files open at the same time: the zip file and the ical file,
    # so we used a nested with block.
    with zipfile.ZipFile(get_zip_file_name()) as zip_file:
        with zip_file.open(get_ical_file_name(zip_file)) as ical_file:        
            write_output(get_desired_events(ical_file, desired_date), desired_date)

if __name__ == '__main__':
  main()
