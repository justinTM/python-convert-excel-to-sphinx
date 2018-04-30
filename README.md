# python-convert-excel-to-sphinx
A python script to convert Excel files into reStructuredText-formatted documentation  

![Screenshot: convert Excel workbook to Sphinx / reST documentation](https://github.com/justinTM/python-convert-excel-to-sphinx/raw/master/excel_to_sphinx_screenshot.png)

## Template files
In order to extract data from an Excel sheet and create a reST-formatted document, template files are used. Inside these template files will be a mixture of reST-formatted text and tags used by the python script to populate the template (eg. `{loop_start}`.

Example snippet from `example_sheet1_template.txt` file:
```
{loop_start}

{Data Tag}
------------------------------------------------------------
{Description}
  :Field Type: {Field Type}
  :Units: {Units}
  :Source: {Source}
  :Logic: {Logic}

{loop_end}
```

  ### Column header tags
  In the above snippet, `{Data Tag}` is an example column header in the Excel sheet "test_worksheet.xlsx", as are `{Field Type}
`, `{Units}`. `{Source}`, and `{}`. The python script will look for any tags matching a column header (inside brackets { } ) and substitute the actual cell contents.

  ### Loop tags
  To populate ALL rows from the Excel sheet into the template file, a loop is specified by the tags `{loop_start}` and `{loop_end}`. The script will duplicate the lines between these two tags -- one duplicate for each row in a sheet. Then, row-by-row, the script will replace tags matching each column header with actual cell values.

## Functions
These are the functions found inside populate_templates_from_sheets.py

* ### get_indent_of_string(string, string_to_find)
  Returns a string representing the indent of a string found within a string.
  
* ### add_indent_to_string(string, indent_string, start_line_index=0, num_lines=0)
  Adds an indent string before each line. Default num_lines=0 is all lines. indentation_string eg. '\t\t' or '    '

* ### get_string_between_strings(s, s_begin, s_end)
  Returns the string found between two strings, eg. '{start}}' and '{end}'

* ### sanitize(string)
  Prevents characters in a string from being interpreted as reST formatting.
  
* ### get_cell_values_from_sheet(xlrd_sheet)
  Returns a 2D list (list[rows][columns]) of cell values from an Excel sheet.

* ### make_data_csv_string(cell_values)
  Returns a comma-separated string of cell values, given 2D list of cells (eg.   '"cell1A", "cell1B"\n#    \t\t"cell2A", "cell2B"')  
    Indents each line to keep reST formatting for csv tables.

* ### make_headers_csv_string(cell_values)
  Returns a comma-separated string of header names, given 2D list of cells.  
    eg. '"header1", "header2", "header3"'

* ### populate_sphinx_template(cell_values, template_filename, output_filename)
  Replaces data tags (eg. "{Description}") in template file with Excel values.



