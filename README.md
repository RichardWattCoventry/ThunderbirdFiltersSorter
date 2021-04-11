# ThunderbirdFiltersSorter

VBScript utility to automate sorting the message filters I use with Mozilla Thunderbird into alphabetical order for easier processing.

This utility is comprised of 2 files:

1) OptimiseMessageFilters.cmd - this is a batch file that makes a backup copy of the current Thunderbird message filters file before launching the sorting script.
2) MessageFilters.vbs - this is the script that performs the sorting process using a quick sort method on an index created from the file, and then uses that to write the sorted list back to the file.

This utility has to ensure that it does not affect the first 2 lines of the file, which are the version number and logging flag lines, and it will be updated over time with a view to convert it from a batch file and script combination into a single Windows executable.
