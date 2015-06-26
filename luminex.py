# -*- coding: utf-8 -*-

#-------------Modules----------------------------------------------------------
from collections import defaultdict
from os import path
from csv import DictReader, writer, QUOTE_ALL
from glob import glob
from math import log
from re import findall, I
from sys import argv
from xlrd import open_workbook
import xlwt


#-------------Function to read files and perform needed oprations--------------
def process(
    sheet_name,
    column_starts,
    column_ends,
    row_starts,
    sheet_name2,
    factor_number,
    sheet_name3
):
    # Static variable initializaitons
    items = []
    mon_values = []
    factors = []
    # Builtin function to retrive all csv files from folder
    for csvfiles in glob(path.join('Before_Calculation', '*.*')):
        # Builtin function to open file
        workbook = open_workbook('%(file_name)s' % {'file_name': csvfiles})
        # Method used to read data from sheet
        sheet = workbook.sheet_by_name(
            '%(sheet_name)s' % {'sheet_name': sheet_name}
        )
        sheet2 = workbook.sheet_by_name(
            '%(sheet_name)s' % {'sheet_name': sheet_name2}
        )
        sheet3 = workbook.sheet_by_name(
            '%(sheet_name)s' % {'sheet_name': sheet_name3}
        )
        #------Looping-----------
        for row, plate in zip(
            range(sheet.nrows)[int(row_starts):],
            ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        ):
            #------Looping-----------
            for col in range(sheet.ncols)[int(column_starts):int(column_ends)]:
                luminex_values = sheet.cell_value(row, col)
                # regular expression to find all '\n'
                next_line = findall('\n', luminex_values, flags=I)
                # Variable initialization
                data_sets = ''
                # Condition to check string in next_line or not
                if next_line:
                    # Split by new line
                    data_sets =  luminex_values.split(next_line[0])
                    # assigning values to varibles
                    well = plate + str(col)
                    construct = data_sets[0]
                    pedigree = data_sets[1]
                    tissue = data_sets[3].split(' ')[0].replace(',', '')
                    # Checking data_set[3] for alphabet, numeric or both
                    if ',' in data_sets[3]:
                        allocation = '' if data_sets[3].split(
                            ','
                        )[-1].isalpha() else data_sets[3].split(',')[-1]
                    else:
                        allocation = '' if data_sets[3].split(
                            ' '
                        )[-1].isalpha() else data_sets[3].split(' ')[-1]
                    stage = data_sets[4].split('<')[0].replace(' ', '')
                    # Appending values to list called items
                    items.append([
                        well, construct, pedigree, tissue, allocation, stage
                    ])
                # Else executes when next_line is empty
                #else:
                #    well = plate + str(col)
                #    items.append([well, '', '', '', '', ''])
        #------Looping-----------
        for row in range(sheet2.nrows)[1:]:
            # Variable Initializations
            col_values = []
            substract_values = []
            col_header = []
            #------Looping-----------
            for col in range(sheet2.ncols):
                # Appending data
                col_header.append(sheet2.cell_value(0, col))
                col_values.append(sheet2.cell_value(row, col))
            #------Looping-----------
            for col in range(sheet3.ncols):
                # Bit of calculation to get expected results
                substract_values.append(
                    sheet2.cell_value(row, col) - sheet3.cell_value(0, col)
                )
            mon_values.append(col_values + substract_values)
            factors.append(substract_values)
        # Dump data into file
        with open(
            'After_Calculation/new_calculated_values.csv' , 'w'
        ) as resource:
            writer(
                resource,
                delimiter=',',
                doublequote=True,
                lineterminator='\n',
                quoting=QUOTE_ALL,
                skipinitialspace=False,
            ).writerow(
                [
                    'Well_Position',
                    'Construct',
                    'Pedigree',
                    'Tissue',
                    'Allocation',
                    'Stage'
                ] + col_header + col_header + [
                    'Normalization Factor'
                ] + col_header + col_header
            )
    # Static variable initializaitons
    total = 0
    mon_1 = []
    #------Looping-----------
    for factor in factors:
        total += factor[int(factor_number)]
    total = total / len(mon_values)
    #------Looping-----------
    for factor in factors:
        # Variable initializaiton
        normalizations = []
        log_values = []
        # Bit of calculation to get expected results
        normalization = total / factor[int(factor_number)]
        normalizations.append(normalization)
        #------Looping-----------
        for factor_value in factor:
            normalizations.append(factor_value * normalization)
        #------Looping-----------
        for factor_value in factor:
            log_value = factor_value * normalization
            try:
                log_values.append(log(log_value, 2))
            except:
                log_values.append(log(-1 * log_value, 2))
        mon_1.append(normalizations + log_values)
    # Dump data into file
    with open('After_Calculation/new_calculated_values.csv', 'a') as resource:
        writer(
                resource,
                delimiter=',',
                doublequote=True,
                lineterminator='\n',
                quoting=QUOTE_ALL,
                skipinitialspace=False,
            ).writerows(
                item + value_ + mon for item, value_, mon in zip(
                    items, mon_values, mon_1
                )
            )

# argv[1]: sheet name, argv[2]: column Start, argv[3]: Column end,
# argv[4]: Row starts from, argv[5]: sheet name2, argv[6]: total
# number of columns from sheet2, argv[7]: sheet name3

#--------start here------------------------------------------------------------
process(argv[1], argv[2], argv[3], argv[4], argv[5], argv[6], argv[7])