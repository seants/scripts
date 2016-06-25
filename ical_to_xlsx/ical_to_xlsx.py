import arrow
from dateutil import tz
from ics import Calendar
import openpyxl
import os
import os.path
from stat import ST_CTIME
import sys
import zipfile

ZIP_DIRECTORY = 'C:\\Users\\Sean\\Downloads'
ZIP_STR = '.zip'
ICAL_STR = '.ical'
OUT_FILE_TYPE = '.xlsx'

def get_desired_date():
    if len(sys.argv) == 1:
        return arrow.now().replace(months=-1)
    elif len(sys.argv) == 3:
        year = int(sys.argv[1])
        month = int(sys.argv[2])
        if year < 1970 or year > 3000 or month < 1 or month > 12:
            raise Exception("Unrecognized year and month: %i %i; expected date like 2016 5" % (year, month))
        return arrow.get(year, month, 1, tzinfo=tz.tzlocal())
    else:
        raise Exception('Unrecognized number of command-line arguments: %i; expected 0 or 2' % (len(sys.argv) - 1))

def get_zip_file_name():
    files = os.listdir(ZIP_DIRECTORY)
    candidate_files = [os.path.join(ZIP_DIRECTORY, filename) for filename in files if ZIP_STR in filename and ICAL_STR in filename]
    if len(candidate_files) == 0:
        raise Exception('No .ical ZIP file found. Try downloading again, and make sure you are in the right directory: %s.' % ZIP_DIRECTORY)
    entries = [(os.stat(name), name) for name in candidate_files]
    entries.sort(key=lambda tuple: tuple[0][ST_CTIME], reverse=True)
    return entries[0][1]

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

def write_output(events, name):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    for i, event in enumerate(events):
        worksheet.append([event.name, event.begin.format('MMMM DD'), event.duration])
        worksheet["D%i" % (i + 1)].number_format
    workbook.save(name)

def get_desired_events(ical_file, date):
    calendar = Calendar(ical_file.read().decode('iso-8859-1'))
    events = filter_by_date(calendar.events, date)
    events = filter_by_custom(events)
    strip_number(events)
    return events

def get_output_name(date):
    name = date.format('YYYY MM') + OUT_FILE_TYPE
    files = os.listdir('.')
    if name in files:
        raise Exception('File %s already exists in current directory. Please delete.' % name)
    return name

def main():
    desired_date = get_desired_date()
    output_name = get_output_name(desired_date)

    # We need two files open at the same time: the zip file and the ical file,
    # so we used a nested with block.
    with zipfile.ZipFile(get_zip_file_name()) as zip_file:
        with zip_file.open(get_ical_file_name(zip_file)) as ical_file:        
            write_output(get_desired_events(ical_file, desired_date), output_name)

if __name__ == '__main__':
  main()
