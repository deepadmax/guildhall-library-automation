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
[bright_yellow]
[b][ TIPS ][/b]

* You can click and drag a file onto the input field to copy its path.

* Leave the input field empty to use the default value.

* If a full path is not given, it will be relative to wherefrom this program is executed.
[/bright_yellow]""".lstrip())


def input_path(prefix='', default=None, check=True):
    """ Wait for user to input a valid path to file """

    while True:
        if default:
            print(f"Default: '{default}' {prefix}", end='')

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
            print('[blue]File found[/blue]\n')
            return path
        
        if not check:
            return path
        
        print(f"\n[bright_red]Either '[b]{path}[/b]' is not a file or the path does not exist.[/bright_red]")
        print('Please try another path.\n')


def print_busy(activity):
    print(f'[bright_cyan]{activity}...[/bright_cyan]')


# Input in which file the list of people and their IDs are located
print("""
[b]Where is your Names & IDs spreadsheet located?[/b]
[i]
- Recommended file format: [bright_blue]XLSX (Excel Spreadsheet Extended)[/bright_blue]
- Table should have two columns, [bright_blue]full name[/bright_blue] & [bright_blue]library ID[/bright_blue], with headers included.
- Full names must be formatted just as they are in the report.
  Example: 'Wimbleton, Rose (Ms)'
[/i]""")
USERS_PATH = input_path(default='examples/users.xlsx')


# Input which file to extract entries from
print("""
[b]Where is your user report located?[/b]
[i]
- Recommended file format: [bright_blue]TXT (Plain Text Document)[/bright_blue]
[/i]""")
REPORT_PATH = input_path(default='examples/report.txt')


# Input where to save the output data
print("""
[b]Where would you like to save the output file?[/b]
[i]
- Make sure the directory already exists.
- Filename should be suffixed with [b]'.xlsx'[/b]
[/i]""")
OUTPUT_PATH = input_path(default='output.xlsx', check=False)


# Load spreadsheet of names & IDs to match
print_busy('\n\nReading names & IDs from spreadsheet')

df = pd.read_excel(USERS_PATH)
df = df.to_records(index=False)

id_names_lookup = {
    _id: name
    for name, _id in df
}


# Extract users with overdue entries in report
print_busy('Extracting relevant information from user report')

overdue_ids = set()

with open(REPORT_PATH) as f:
    text = f.read()
    match = re.findall('\s*(.+, .+)[\n\s]+id:([\w\d-]+)\s+((?:.|\n)+?Charges)', text)

    for name, _id, extra in match:
        if extra.count('reason:OVERDUE'):
            overdue_ids.add(int(_id))


# Find which IDs exist in both sets
print_busy('Comparing users to overdue entries')

overlap = overdue_ids.intersection(id_names_lookup.keys())
overdue_entries = [
    (id_names_lookup[_id], _id)
    for _id in overlap
]


# Save result to spreadsheet
print_busy('Saving to Excel Spreadsheet')

df = pd.DataFrame(overdue_entries, columns=['Name', 'ID'])
df.to_excel(OUTPUT_PATH, index=None, header=True)


# Success message
print('\nSuccessfully generated table!')
input() # Leave window open