import arrow
import csv
from dateutil import tz
from ics import Calendar
import openpyxl
import os
import os.path
import zipfile

# TODO: fail immediately if filename already exists

ZIP_DIRECTORY = 'E:\\Downloads'
ZIP_ICAL_STR = '.ical.zip'
OUT_FILE_TYPE = '.xlsx'

SECONDS_PER_MINUTE = 60
MINUTES_PER_HOUR = 60

COST_PER_HOUR = 22

def filter_by_date(events, date):
    return events[date.floor('month'):date.ceil('month')]

def filter_by_custom(events):
    return [event for event in events if event.name.lower() != 'eva out']
    
def strip_number(events):
    for event in events:
        event.name = ''.join(filter(lambda a: not a.isnumeric(), event.name)).strip()

l = os.listdir(ZIP_DIRECTORY)

def calculate_cost(event):
    return event.duration.seconds / SECONDS_PER_MINUTE / MINUTES_PER_HOUR * COST_PER_HOUR

def write_csv(events, date):
    name = date.format('YYYY MMMM') + OUT_FILE_TYPE
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    for i, event in enumerate(events):
        worksheet.append([event.name, event.begin.format('MMMM DD'), event.duration, calculate_cost(event)])
        worksheet["D%i" % (i + 1)].number_format
    worksheet.append(["", "", "=SUM(C1:C%i)" % len(events), "=SUM(D1:D%i)" % len(events)])
    workbook.save(name)

candidate_files = [filename for filename in l if ZIP_ICAL_STR in filename]
if len(candidate_files) == 1:
    zip_file_name = candidate_files[0]
elif len(candidate_files) == 0:
    raise Exception('No .ical ZIP file found. Try downloading again, and make sure you are in the right directory.')
else:
    raise Exception('Multiple .ical ZIP files found. Please delete all old .ical ZIP files.')

with zipfile.ZipFile(os.path.join(ZIP_DIRECTORY, zip_file_name)) as zip_file:
    ical_file_names = zip_file.namelist()

    if len(ical_file_names) != 1:
        raise Exception('ZIP archive had %i files; expected 1.' % len(ical_file_names))
        
    with zip_file.open(ical_file_names[0]) as ical_file:        
        c = Calendar(ical_file.read().decode('iso-8859-1'))
        desired_date = arrow.now().replace(months=-1)
        # desired_date = arrow.get(2016, 5, 1, tzinfo=tz.tzlocal())
        events = filter_by_date(c.events, desired_date)
        events = filter_by_custom(events)
        strip_number(events)
        write_csv(events, desired_date)
