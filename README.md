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

        $: python at12xl.py path/to/my/xml/file.xml

or 

        $: python at12xl.py path/to/my/xml/file.xml -e newname.xls

You can adjust the #! statement at the begining of the script to point to your Python binary so you don't have to explicitly call Python

        $: ./ati2xl.py path/to/my/xml/file.xml -e newname.xls

Missing Functionality
---------------------

This code does not extract the following sections of the Atlas.ti XML dump:

* codeLinkProtos
* codeLinks
* hyperLinkProtos
* memoMemoLinks
* codeMemoLinks