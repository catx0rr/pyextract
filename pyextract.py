#!/usr/bin/env python

import pandas as pd
import csv
import xlsxwriter

from statuslib import status as s

### Change the configuration here.

filename = 'your_file.xlsx'             # make sure to add file extension
                                        # e.g. (.xls, .xlsx)
output_filename = 'desired_filename'    # without file extension
needed_data = [1, 2, 3]                 # change int to needed rows.


### Do not touch the code below

def excel_to_csv(filename, name='data'):
    
    out_name = '%s%s' % (name, '.csv')

    df = pd.read_excel(filename, sheet_name=0)
    df.to_csv(out_name, index=None, header=True)
    
    s.success('Data saved on current directory Filename: %s' % (out_name))
    return out_name


def get_raw_data(filename, row=1):

    # Fetch the row raw data in a list

    # Align with index 0
    row -= 1

    with open(filename) as f:
        reader = csv.reader(f)
        
        data_list = []
        s.success('Raw data extracted.')

        for index, data in enumerate(reader):
            if index == row:
                data_list.append(('\n'.join(data).splitlines()))
 
        return data_list


def xlsx_write(data_list, output_name):

    output_name = ('%s%s' % (output_name, '.xlsx'))

    with xlsxwriter.Workbook(output_name) as f:
        worksheet = f.add_worksheet()

        s.working('Writing data to %s' % (output_name))
        for index, data in enumerate(data_list):
        
            for item in data:
                worksheet.write_row(index, 0, item)

        s.success('Done writing data.')

 
def fetch_data(filename, index_list):

    output_data = []
    for index in index_list:
        raw_data = get_raw_data(filename, row=index)
        output_data.append(raw_data)

    return output_data


def print_rows(excel):

    with open(excel) as f:
        reader = csv.reader(f)
        
        for row in reader:
            join = ' '.join(row)
            
            print('%s%s' % (join, '\n'))



def main(filename, output_name='generated_data', needed_data):
    
    name = excel_to_csv(filename, 'mydata')

    fetched_data = fetch_data(name, needed_data)
    
    xlsx_write(fetched_data, output_name)


if __name__ == '__main__':

    main(filename, output_filename, needed_data)
