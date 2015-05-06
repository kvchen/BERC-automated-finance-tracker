BERC Automated Finance Tracker
==============================

## Instructions

### Installation

This program must be run from a Windows computer with Microsoft Excel 
installed. The following instructions are for Python 3.4.3, the latest stable
release as of this documentation - however, it should work with any Python
version >= 3.4, or any Python version >= 3.0 with Pip installed.

1. Download Python 3.4.3 [here](https://www.python.org/downloads/release/python-343/) at the bottom. Make sure to choose the `Windows x86-64 MSI installer`.
2. Begin the installation. Make sure that during the setup, "add Python to PATH" is selected to be installed to disk!
![add Python to Path](http://i.imgur.com/m4zyF7v.png)
3. Download PyWin32 [here](http://sourceforge.net/projects/pywin32/files/pywin32/Build%20219/). Make sure to pick the x64 version corresponding to the Python version you installed in the previous steps (for this README, `pywin32-219.win-amd64-py3.4.exe`). This library allows Python to interact with Windows applications.
4. In the main folder `BERC-automated-finance-tracker`, right-click `get-dependencies.bat` and select "Run as administrator." This allows Python to download a few libraries that it needs to interact with PayPal and Google Sheets.
5. Check the settings in `config.yml` (right-click and open with Notepad) and make sure everything is up-to-date. If you are running this on behalf of BERC, please contact the current maintainer and request a copy.

### GUI Usage

The program can be started from `run.pyw` in the main folder. Upon starting the program, a window should pop up showing three options - "Update," "Backup," and "Exit." Select the correct one and let the program do its thing!


## FAQ

* I see duplicate entries in the Payables tab. How do I get rid of them?
    * To fix this, you need to find the difference between the two entries (the difference might be small). Make sure the entry in the Google reimbursement form matches up exactly with that in the Excel spreadsheet.

### Modification

This project is created and maintained by [Kevin Chen](https://github.com/kvchen). If modification is necessary, please get in touch with the current maintainer.

This code is free and open source under the [MIT license](https://raw.githubusercontent.com/kvchen/BERC-automated-finance-tracker/master/LICENSE).


