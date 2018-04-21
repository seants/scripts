"""Script to convert ical to xlsx file."""

import os
import os.path
from stat import ST_CTIME
import sys
import zipfile

import arrow
from dateutil import tz
from ics import Calendar
import openpyxl

# Directory where the .ical.zip file has been downloaded
ZIP_DIRECTORY = 'C:\\Users\\Sean\\Downloads'

ZIP_STR = '.zip'
ICAL_STR = '.ical'
OUT_FILE_TYPE = '.xlsx'

TIME_RANGE = 'month'

def get_desired_date():
    """Returns a date within the desired range."""
    if len(sys.argv) == 1:
        return arrow.now().replace(months=-1)
    elif len(sys.argv) == 3:
        year = int(sys.argv[1])
        month = int(sys.argv[2])
        if year < 1970 or year > 3000 or month < 1 or month > 12:
            raise Exception(
                "Unrecognized year and month: %i %i; expected date like 2016 5."
                % (year, month)
            )
        return arrow.get(year, month, 1, tzinfo=tz.tzlocal())
    else:
        raise Exception(
            "Unrecognized number of command-line arguments: %i; expected 0 or 2."
            % (len(sys.argv) - 1)
        )

def get_zip_file_name():
    """Gets the full name of the zip file."""
    files = os.listdir(ZIP_DIRECTORY)
    candidate_files = [os.path.join(ZIP_DIRECTORY, filename)
                       for filename in files
                       if ZIP_STR in filename and ICAL_STR in filename
                      ]
    if len(candidate_files) == 0:
        raise Exception(
            "No .ical ZIP file found. Try downloading again, and make sure"
            + "you are in the right directory: %s." % ZIP_DIRECTORY
        )
    entries = [(os.stat(name), name) for name in candidate_files]
    entries.sort(key=lambda tuple: tuple[0][ST_CTIME], reverse=True)
    return entries[0][1]

def get_ical_file_name(zip_file):
    """Gets the name of the ical file within the zip file."""
    ical_file_names = zip_file.namelist()
    if len(ical_file_names) != 1:
        raise Exception(
            "ZIP archive had %i files; expected 1."
            % len(ical_file_names)
        )
    return ical_file_names[0]

def filter_by_date(events, date):
    """Returns only the events that fall within the desired time range."""
    return events[date.floor(TIME_RANGE):date.ceil(TIME_RANGE)]

def write_output(events, name):
    """Writes the events to the output file."""
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    for event in events:
        worksheet.append(
            [event.name, event.begin.format('MMMM DD'), event.duration]
        )
    workbook.save(name)

def get_desired_events(ical_file, date):
    """Returns the events, given the input file."""
    calendar = Calendar(ical_file.read().decode('iso-8859-1'))
    return filter_by_date(calendar.events, date)

def get_output_name(date):
    """Gives the name of the output file, and checks for duplication."""
    name = date.format('YYYY MM') + OUT_FILE_TYPE
    files = os.listdir('.')
    if name in files:
        raise Exception(
            'File %s already exists in current directory. Please delete.'
            % name
        )
    return name

def main():
    """Main script."""
    desired_date = get_desired_date()
    output_name = get_output_name(desired_date)

    # We need two files open at the same time: the zip file and the ical file,
    # so we used a nested with block.
    with zipfile.ZipFile(get_zip_file_name()) as zip_file:
        with zip_file.open(get_ical_file_name(zip_file)) as ical_file:
            write_output(
                get_desired_events(ical_file, desired_date),
                output_name
            )

if __name__ == '__main__':
    main()
