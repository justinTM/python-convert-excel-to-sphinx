#!/usr/bin/env python

import xlrd
import textwrap


# Returns a string representing the indentation of the line
# this is neccessary to properly format the CSV table
def get_indentation_of_line(line):
    # Get the indent as a string (eg. '\t\t' or '   ')
    indentation_string = re.match(r"\s*", line).group()


# Returns a comma-separated string of header names, given 2D list of cells
# eg. '"header1", "header2", "header3"'
def make_headers_csv_string(cell_values):
    headers = ""
    header_row = 0
    num_columns = len(cell_values[header_row])
    # Loop through rows and columns of cell_values
    for cx, column in enumerate(cell_values[header_row]):
        header = cell_values[header_row][cx]
        # Wrap each header with quotations (eg. "header1")
        headers += '"' + header + '"'
        # Add a comma between each column, except the last column
        if cx+1 < num_columns:
            headers += ', '

    return headers


# Returns a comma-separated string of cell values, given 2D list of cells
# eg.   '"cell1A", "cell1B"\n
#    \t\t"cell2A", "cell2B"'
# prepends each line with a string representing the indent (ie. tabs, spaces)
def make_data_csv_string(cell_values):
    csv_string = ""
    header_row = 0
    num_rows = len(cell_values)
    # Loop through rows and columns of cell_values
    for rx, row in enumerate(cell_values):
        # Skip the header row
        if rx <= 0: continue
        for cx, column in enumerate(row):
            cell_value = sanitize(cell_values[rx][cx])
            # Cut long strings down to reduce table size
            if cell_values[0][cx] != "Data Tag":
                cell_value = cell_value #textwrap.wrap(cell_value, 30)[0]
            # Make each Data Tag a link to the corresponding section
            else: cell_value = '`'+cell_value+'`_'
            # Wrap each cell with quotations (eg. "<cell>")
            csv_string += '"' + cell_value + '"'
            # Add commas between each column, except for the last column
            if cx+1 < len(cell_values[0]):
                csv_string += ', '
        # Add a newline after each row, except the last row
        if rx+1 < num_rows:
            csv_string += '\n'

    return add_indent_to_string(csv_string, '    ')


# default lines=0 is all lines, indentation_string eg. '\t\t' or '    '
def add_indent_to_string(string, indent_string,
                         start_line_index=0, num_lines=0):
    # Make an array of each line in string, keeping '\n' for each
    lines = string.splitlines(1)

    # ending line is num_lines if specified, otherwise all lines (ie len(list))
    if num_lines: end_line_num = start_line_index + num_lines
    else: end_line_index = len(lines)

    # Add the indent string before each line
    for i in range(start_line_index, end_line_index):
        lines[i] = indent_string + lines[i]

    # Combine the array back into a single string
    return "".join(lines)


# Returns a 2D list (list[rows][columns]) of cell values from an Excel sheet
def get_cell_values_from_sheet(xlrd_sheet):
    num_rows = xlrd_sheet.nrows
    num_columns = xlrd_sheet.ncols
    if num_rows < 1 or num_columns < 1:
        print("number of rows or columns equals zero,\
        exiting get_cell_values_from_sheet()")
        return

    # Pre-populate the 2D list, to make it obvious if a value was not copied
    cell_values = [
        ["Undefined" for x in range(num_columns)]
         for y in range(num_rows)]

    # Loop through all cells. Enumerate() creates easy indeces 'rx' and 'cx'
    for rx, row in enumerate(cell_values):
        for cx, column in enumerate(row):
            cell = xlrd_sheet.cell(rx, cx)
            if not cell: continue
            # ctype == 0 means empty string, 5 means excel error code
            if (cell.ctype > 0 and cell.ctype < 5):
                cell_values[rx][cx] = str(cell.value)
            else:
                cell_values[rx][cx] = "error: " + str(cell.ctype)

    return cell_values


# Returns the string found between eg. '{start}}' and '{end}'
def get_string_between_strings(s, s_begin, s_end):
    begin_index = s.find(s_begin)
    if begin_index < 0: return ''
    s_middle = s[begin_index:len(s)]
    end_index = begin_index + s_middle.find(s_end)
    string_between = s[begin_index+len(s_begin):end_index]
    return string_between


# Returns a string representing the indent of a string found within a string
def get_indent_of_string(string, string_to_find):
    indent_index = 0
    for line in string.splitlines():
        if string_to_find in line:
            indent_index = line.find(string_to_find)
    indent = ' '*indent_index

    return indent


# Replaces data tags (eg. "{Description}") in template file with Excel values
def populate_sphinx_template(cell_values, template_filename, output_filename):
    # Read in the template file
    with open(template_filename, 'r') as file :
        filedata = file.read()

    # Paste the csv table string (headers & rows have separate requirements)
    data_csv_string = make_data_csv_string(cell_values)
    headers_csv_string = make_headers_csv_string(cell_values)
    filedata = filedata.replace('{headers}', headers_csv_string)
    filedata = filedata.replace('{csv_table}', data_csv_string)

    # Select string between start and end tags (ie. the loop template)
    loop_template = get_string_between_strings(
                        filedata, '{loop_start}', '{loop_end}')

    # Duplicate the loop template for all rows, skipping header row
    looped_string = ""
    for i in range(1, len(cell_values)):
        looped_string += '\n'+loop_template

    # Replace each header tag (eg. {Description}) with the cell value
    for rx, row in enumerate(cell_values):
        if rx == 0: continue
        for cx, column in enumerate(row):
            header = cell_values[0][cx]
            value = cell_values[rx][cx].replace('\n', '\n')
            header_to_search = '{'+header+'}'
            indent = get_indent_of_string(looped_string, header_to_search)
            value = add_indent_to_string(value, indent, 1)
            # Replace only the first occurence
            looped_string = looped_string.replace(header_to_search, value, 1)

    # Replace the looped data, and delete the loop tags
    filedata = filedata.replace(loop_template, looped_string)
    filedata = filedata.replace('{loop_start}', '')
    filedata = filedata.replace('{loop_end}', '')

    # Write the file out again
    with open(output_filename, 'w', encoding='utf_8') as file:
        file.write(filedata)

    return


# Prevent characters in a string from being interpreted as reST formatting
def sanitize(string):
    string = string.replace('"', '\\"')
    string = string.replace('"', '\\“')
    string = string.replace('"', '\\”')
    string = string.replace(',', '\\,')
    #string = string.replace('\n', '|\n')
    return string


def main():
    print('hello world!')

    # Open an Excel workbook and get the first sheet
    print('Opening workbook...')
    book = xlrd.open_workbook('test_workbook_new.xlsx',
                              encoding_override='cp1252')

    databases = ['raw_spc_records', 'spc_records', 'drill_parameters']
    for database in databases:
        sheet = book.sheet_by_name(database)
        # Convert the excel table to a 2D list in python
        cell_values = get_cell_values_from_sheet(sheet)
        # Convert 2D list to reStructuredText and paste into the template
        template_filename = database+'_template.txt'
        output_filename = database+'.rst'
        populate_sphinx_template(cell_values,
                                 template_filename,
                                 output_filename)

    print('Done!')
if __name__ == '__main__':
    main()
