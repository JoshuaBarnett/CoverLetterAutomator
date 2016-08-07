# CoverLetterAutomator
Python script to generate customized cover letters for a batch of job listings, given a template

Requires .xls or .xlsx spreadsheet and .docx template
Dependencies: openpyxl and python-docx

Not tested on Windows but should work fine?

Use addReplacementList/3 with a key and a start location.
  Key = string in Cover Letter template to be replaced
  The start location is a row and column in your spreadsheet; this should correspond with the start of a contiguous vertical list. See example for details
