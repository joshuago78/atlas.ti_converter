atlas.ti_converter
==================

Converts an Atlas.ti XML dump to an Excel spreadsheet

Requirements
-----------
* Python 2.7 (not tested with Python 3)
* lxml Python library for parsing xml
* xlwt Python library for writing Excel files

Usage
-----

Simply call the python script and supply the arguments for the input and output files

The path to the input XML file is required.

The filename of the output Excel file is optional (defaults to same name as xml with different extension).  Denote it with the flag -e or --excel.

Example:

        $: ./ati2xl.py path/to/my/xml/file.xml -e newname.xls
