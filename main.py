from PIL import Image
from colormap import rgb2hex
import xlsxwriter

import numpy as np

from pprint import pprint

import pandas as pd
import os

import subprocess as sub

from xlsxwriter.utility import xl_col_to_name
from xlsxwriter.utility import xl_rowcol_to_cell

import time


def format_cells_blank(cells_to_format=None,
                       bg_color=None, workbook=None, sheet=None):
    formatt = workbook.add_format()
    formatt.set_bg_color(bg_color)

    sheet.conditional_format(
        xl_rowcol_to_cell(cells_to_format[0], cells_to_format[1]), {
            'type': 'blanks',
            'format': formatt
        })


start_time = time.time()
filename = 'img1.png'
img = Image.open(filename)

w, h = img.size
print(w, h)

img_RGB = img.convert('RGB')

output_filename, _ = os.path.splitext(filename)
output_filename += ".xlsx"

with pd.ExcelWriter(output_filename, mode='w', engine='xlsxwriter') as writer:
    sheet_name = "Image"
    workbook = writer.book
    sheet = workbook.add_worksheet()

    for i in range(w):
        sheet.set_row(i, 11)
    for i in range(h):
        sheet.set_column(f"{xl_col_to_name(i)}:{xl_col_to_name(i)}", 1.5)

    for i in range(w):
        print(i)
        for j in range(h):
            r, g, b = img_RGB.getpixel((j, i))
            format_cells_blank(
                (i, j), rgb2hex(r, g, b), workbook, sheet)
print(f"--- {time.time() - start_time} seconds ---")
sub.Popen(output_filename, shell=True)
