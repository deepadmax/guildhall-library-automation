import os
import re

from pathlib import Path

from rich import print
from rich.table import Table

import pandas as pd


print("""
┌──────────────────────┐
│ FIND OVERDUE ENTRIES │
└──────────────────────┘
[gold1]
[b][ TIPS ][/b]

* You can click and drag a file onto the input field to copy its path.

* Leave the input field empty to use the default value.

* If a full path is not given, it will be relative to wherefrom this program is executed.
[/gold1]""".lstrip())


def input_path(prefix='', default=None, check=True):
    """ Wait for user to input a valid path to file """

    while True:
        if default:
            print(f"Default: [spring_green3]'{default}'[/spring_green3] {prefix}", end='')

        path = input('> ')

        if path == '':
            path = default


        # Replace backslashes with forward slashes
        path = path.replace('\\', '/')

        # Strip of redundant spaces
        path = path.strip()
        
        # Remove surrounding quotes
        if f'{path[0]}{path[-1]}' in ['""', "''"]:
            path = path[1:-1]
        

        path = Path(path)

        if not check:
            return path

        if os.path.isfile(path):
            print('[i]File found.[/i]\n')
            return path
        
        if not check:
            return path
        
        print(f"\n[bright_red]Either '[b]{path}[/b]' is not a file or the path does not exist.[/bright_red]")
        print('Please try another path.\n')


def print_busy(activity):
    print(f'[deep_sky_blue1]{activity}...[/deep_sky_blue1]')

def failure(string):
    print(f'[i][bright_red]Error: {string}[/bright_red][/i]')
    input()
    exit()


# Input in which file the list of people and their IDs are located
print("""
[b]Where is your ID spreadsheet located?[/b]
[i]
- Recommended file format: [deep_sky_blue1]XLSX (Excel Spreadsheet Extended)[/deep_sky_blue1]
- Table must have a column labelled [deep_sky_blue1]ID[/deep_sky_blue1], under which the library IDs of all users whom you wish to check are listed.
[/i]""")
USERS_PATH = input_path(default='examples/users.xlsx')


# Input which file to extract entries from
print("""
[b]Where is your user report located?[/b]
[i]
- Recommended file format: [deep_sky_blue1]TXT (Plain Text Document)[/deep_sky_blue1]
[/i]""")
REPORT_PATH = input_path(default='examples/report.txt')


# Input where to save the output data
print("""
[b]Where would you like to save the output file?[/b]
[i]
- Make sure the directory already exists.
- Filename should be suffixed with [b][spring_green3]'.xlsx'[/spring_green3][/b]
[/i]""")
OUTPUT_PATH = input_path(default='output.xlsx', check=False)


# Load spreadsheet of names & IDs to match
print_busy('\n\nReading IDs from spreadsheet')

df = pd.read_excel(USERS_PATH)

# Find column labeled ID with any capitalization
if 'ID' not in df.columns:
    failure('Missing ID column in your spreadsheet.')

lookup_ids = [int(x) for x in df['ID']]


# Extract users with overdue entries in report
print_busy('Extracting relevant information from user report')

report_lookup = {}

with open(REPORT_PATH) as f:
    text = f.read()
    match = re.findall(r'\s*(.+,\s.+)[\n\s]+id:([\w\d-]+)\s+((?:.|\n)+?Charges)', text)

    for name, _id, extra in match:
        report_lookup[int(_id)] = name

report_ids = set(report_lookup.keys())

# Find which IDs exist in both sets
print_busy('Comparing users to overdue entries')

overlap = report_ids.intersection(lookup_ids)
overdue_entries = [
    (report_lookup[_id], _id)
    for _id in overlap
]


# Save result to spreadsheet
print_busy('Saving to Excel Spreadsheet')

df = pd.DataFrame(overdue_entries, columns=['Name', 'ID'])
df.to_excel(OUTPUT_PATH, index=None, header=True)


# Success message
print('\nSuccessfully generated table!')
input() # Leave window open