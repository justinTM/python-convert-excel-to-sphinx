# python-convert-excel-to-sphinx
A python script to convert Excel files into reStructuredText-formatted documentation  

## Functions
These are the functions found inside populate_templates_from_sheets.py

* ### populate_sphinx_template(cell_values, template_filename, output_filename)
  Replaces data tags (eg. "{Description}") in template file with Excel values

* ### get_indent_of_string(string, string_to_find)
  Returns a string representing the indent of a string found within a string

* ### get_string_between_strings(s, s_begin, s_end)
  Returns the string found between two strings, eg. '{start}}' and '{end}'

* ### get_cell_values_from_sheet(xlrd_sheet)
  Returns a 2D list (list[rows][columns]) of cell values from an Excel sheet

* ### add_indent_to_string(string, indent_string, start_line_index=0, num_lines=0)
  Adds an indent string before each line. Default num_lines=0 is all lines. indentation_string eg. '\t\t' or '    '

* ### make_data_csv_string(cell_values)
  Returns a comma-separated string of cell values, given 2D list of cells
    eg.   '"cell1A", "cell1B"\n#    \t\t"cell2A", "cell2B"'
    prepends each line with a string representing the indent (ie. tabs, spaces)

* ### make_headers_csv_string(cell_values)
  Returns a comma-separated string of header names, given 2D list of cells
    eg. '"header1", "header2", "header3"'

* ### sanitize(string)
  Prevents characters in a string from being interpreted as reST formatting

* ### get_indentation_of_line(line)
  Returns a string representing the indentation of the line
