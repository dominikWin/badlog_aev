#!/usr/bin/env python3

import xlrd
import json
import sys

if len(sys.argv) != 2:
    print('Bad number of inputs! use <filename>')
    sys.exit(1)

filename = sys.argv[1]

wb = xlrd.open_workbook(filename)
sheet = wb.sheet_by_index(0)

assert sheet.cell_value(0,0) == 'System Analysis'
assert sheet.cell_value(2,0) == 'Reference Voltage'

ref_voltage = float(sheet.cell_value(2, 1))

data_main_start_x = 6
data_main_start_y = 8
data_main_width = 9

data_raw_start_x = 0
data_raw_start_y = 8
data_raw_width = 5

data = []

row = 0
while True:
    try:
        data_row = []
        # Main Data
        for column in range(data_main_width):
            data_row.append(str(float(
                sheet.cell_value(data_main_start_y+row, data_main_start_x+column))))
        data.append(data_row)

        # Raw Data
        for column in range(data_raw_width):
            data_row.append(str(float(
                sheet.cell_value(data_raw_start_y+row, data_raw_start_x+column))))
        data.append(data_row)
    except:
        break
    row += 1

csv_text = "Time,Current,Voltage,Distance,Position,Speed,Input Power,Incremental Energy,Total Energy,EEProm/Time,EEProm/Current,EEProm/Voltage,EEProm/Marks,EEProm/Relative Marks\n"

for row in range(len(data)):
    row_values = data[row]
    csv_text += ','.join(row_values) + '\n'

bl_values = [{'name': "Reference Voltage", 'value': str(ref_voltage)}, {'name': "Filename", 'value': filename}]

bl_topics = [
    {'name': "Time", 'unit': "s", 'attrs': ['xaxis']},
    {'name': "Current", 'unit': "A", 'attrs': []},
    {'name': "Voltage", 'unit': "V", 'attrs': []},
    {'name': "Distance", 'unit': "m", 'attrs': []},
    {'name': "Position", 'unit': "m", 'attrs': []},
    {'name': "Speed", 'unit': "m/s", 'attrs': []},
    {'name': "Input Power", 'unit': "W", 'attrs': []},
    {'name': "Incremental Energy", 'unit': "J", 'attrs': []},
    {'name': "Total Energy", 'unit': "J", 'attrs': []},
    {'name': "EEProm/Time", 'unit': "ms", 'attrs': []},
    {'name': "EEProm/Current", 'unit': "count", 'attrs': []},
    {'name': "EEProm/Voltage", 'unit': "count", 'attrs': []},
    {'name': "EEProm/Marks", 'unit': "wheel counts", 'attrs': []},
    {'name': "EEProm/Relative Marks", 'unit': "wheel counts", 'attrs': []},
]

bl_header = {'values': bl_values, 'topics': bl_topics}

bl_text = json.dumps(bl_header) + "\n" + csv_text

print(bl_text)