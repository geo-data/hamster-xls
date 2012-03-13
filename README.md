# Hamster-xls

## Overview

This is a small command line python script which creates a .xls (excel
compatible) spreadsheet from the
[Project Hamster Time Tracking](http://projecthamster.wordpress.com/)
application. Because managers like spreadsheets :) It can extract data
for a specified date range (defaulting to the last 30 days), and can
filter the data based on simple search terms.

## Synopsis

From the inbuilt help:

    usage: hamster-xls [-h] [-e DATE] [-s DATE] [-o FILE] [-f FILTER]
    
    Create a .xls spreadsheet from your hamster time tracking data, with the
    ability to specify the date range to extract.
    
    optional arguments:
      -h, --help            show this help message and exit
      -e DATE, --end-date DATE
                            the date tracking data should end at, inclusive in the
                            format YYYY-MM-DD (defaults to today)
      -s DATE, --start-date DATE
                            the date tracking data should start from, inclusive in
                            the format YYYY-MM-DD (defaults to 30 days before the
                            end date)
      -o FILE, --output FILE
                            the .xls spreadsheet file to create (defaults to
                            `./hamster-timesheet.xls`)
      -f FILTER, --filter FILTER
                            limit the data based on search terms. Comma (",")
                            translates to boolean OR and space (" ") to boolean
                            AND.

## Requirements

- [Python](http://www.python.org) > 2.5 < 3
- [Hamster](http://projecthamster.wordpress.com/) to be installed and
  running, as the application connects to the hamster service using
  [D-Bus](http://www.freedesktop.org/wiki/Software/dbus).
- The [xlwt](http://pypi.python.org/pypi/xlwt) python module for
  generating the spreadsheets.
