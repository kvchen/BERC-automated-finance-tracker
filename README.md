BERC Automated Finance Tracker
==============================

### Instructions

##### GUI Usage

The program can be started from `run.py` in the main folder. Upon starting the
program, a terminal and GUI window should open up. The terminal is used only
for debugging purposes, and will log output as the program is running.

While running an operation, the window may look like it is frozen. This is
normal, and you can check on the progress in the terminal.

By default, the GUI will hide all runtime errors. To display them for debugging
purposes, start the program from the command line as described below.

##### Command Line Usage

The updater can also be started from the command line. `run.py` accepts four
arguments:

* `--backup`: creates a copy of the current workbook in a backup directory
* `--update`: attempts to synchronize the workbook information
* `--restore`: restores the most recent backup to the current working directory
* `--test`: runs tests, should only be used by the current maintainer for
debugging purposes

For instance, to run the program and update the workbook:

```
> python run.py --update
```

Errors _will_ be logged in the command prompt when using the command line.


### Installation

This program must be run from a Windows computer with Microsoft Excel (the 
program leverages the win32com api to interact with Microsoft Office).

Under copyright law, I'm not allowed to redistribute some of the installer
executables, so you're going to have to download/install them on your own.

The following instructions are for Python 3.4.1 (The latest stable release as 
of this documentation), but can be extended to any Python version >= 3.4, or 
any Python version >= 3.0 with Pip installed.

1. Download Python 3.4.1 from [this link](https://www.python.org/ftp/python/3.4.1/python-3.4.1.amd64.msi)
and begin installation. Make sure that during the setup, you select 
"add Python to PATH" to be installed to disk!  
![add Python to Path](http://i.imgur.com/m4zyF7v.png)
2. Reboot your computer to let changes take place.
3. Download PyWin32 from [this link](http://sourceforge.net/projects/pywin32/files/pywin32/Build%20219/pywin32-219.win-amd64-py3.4.exe/download) and begin installation. This library allows Python to interact with 
Windows applications (in our case, Excel). Just keep clicking next, and the
installation should proceed smoothly.
4. In the main folder, right-click `lib_installer.bat` and select "Run as 
administrator". This allows Python to download a few libraries that it needs
to interact with PayPal and Google Sheets.
5. Set all relevant setting options in a file called `config.ini` within the 
main program folder. If you are running this on behalf of BERC, please contact
the current maintainer and request a copy.



### Modification

This project was created by [Kevin Chen](https://github.com/kvchen), and is 
currently being maintained by [Kevin Chen](https://github.com/kvchen). If 
modification is necessary, please get in touch with the current maintainer.

This code is open-sourced and may be used and modified within the bounds of 
[the license](https://raw.githubusercontent.com/kvchen/BERC-automated-finance-tracker/master/LICENSE).


