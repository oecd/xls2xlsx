# xls2xlsx

PowerShell script to batch convert old XLS files to more modern XLSX files.

## Description

This script traverses a directory and any subdirectories searching for old `xls` Excel files. It will convert them to `xlsx` files.

> Note that the original files will be stored in a `_xls` directory that will be created. These `_xls` directories are excluded when searching for `xls` files.

## Usage

Create a shortcut of the `bat` file on your Desktop. Then drag and drop a directory onto the shortcut.  You do not have to be Administrator of your machine.  This has been tested with PowerShell version 5.1 but I don't think there are any v5-specific things in the code.

## Limitations

This has been tested (and is intended for) mainly for `xls` Excel files that have been produced by the TOG. It will quite likely work for others as well. However, if there are links that need to be updated or other weird things, it may require human intervention.

## Todo

* [ ] Add more informative output
* [ ] Do more error checking

## Acknowledgement

This script is based on this gist: https://gist.github.com/gabceb/954418
