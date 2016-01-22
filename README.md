Overwrite
=========

Overwrite is a quick and specific use case script for unrestricted bulk 
manipulation of data.

> **WARNING** -
> Overwrite was created for a specific use case scenario.
> One size does not fit all. Pull or deploy at your own risk.

Requirements and Dependencies
-----------------------------
Overwrite was designed for use on .xlsx files.

Software
   * [Python](https://python.org) 2.7 or higher
   * [openpyxl](https://bitbucket.org/openpyxl/openpyxl) 2.4 or higher


Installation and Setup
----------------------
1. Download and install [Python 2.7](https://www.python.org/downloads/)
2. Set environment PATH for python by using the following command line 
```path %path%;C:\Python27```, where C:\Python27 is the path of the python installation
3. Install openpyxl python library by the following command line
```
pip install openpyxl
```


Running
-------
1. Copy files to be overwritten into ```files``` directory, 
ensure there is a master sheet xlsx file with the string 
```Master Sheet``` somewhere in the name.
2. Click on ```overwrite.py``` to run the program or run manually
using the following command
```
python overwrite.py
```


Future Work
-----------
None


License
-------
Copyright 2015-2016 C.I.Djamaludin

Licensed under the GPLv3.

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

