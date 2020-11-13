# Guildhall Library Automation
A few programs to automate things like finding who's overdue and whatnot.
These applications are primarily made to work on Windows, but would anyone request so, they could be adapted where necessary.

Before fiddling with anything, read through and follow the [Setup](#Setup).
Should you not know how to run the individual programs, see [Running an application](#Running_an_application).

## Setup
* To be able to use this kit, you'll need to install the latest version of [Python](https://www.python.org/).
* Regardless of whether you've downloaded this kit for the first time or you've downloaded recent updates to it, always make sure to [install all the dependencies](#Installing_dependencies).

### Installing dependencies
1. Navigate to the top directory of this repository on your computer.
2. Hold Shift and right-click in the file explorer.
3. Select `Open PowerShell Here` or `Open command window here`
4. In the command line, type `pip install -r requirements.txt` and press Enter.


## Applications
### List of all applications
* [Find overdue users](/find_overdue_users)

### Running an application
If your computer has `.py` files associated with Python, they should show up with the Python logo in your file explorer and you should be able to simply double-click the `main.py` file in the directory of application you want to run.

If it does *not* have the `.py` file extension associated with Python, do as follows.

1. Navigate to the directory of the application.
2. Follow step 2 and 3 of [Installing dependencies](#Installing_dependencies).
3. Enter `python main.py` and press Enter.