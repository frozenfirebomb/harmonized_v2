import os
import glob
from datetime import date
import re

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Border, Side

# from file_paths import invnday_path, metal_master_path, steel_codes_path, alum_codes_path, copper_codes_path, save_path, report_path
from chapter_98 import chapter_98_codes


def acquire_chapter_98(harm_code):
    for key, value in chapter_98_codes.items():
        if re.match(key,harm_code):
            return value

print(acquire_chapter_98("8424.90.9080"))