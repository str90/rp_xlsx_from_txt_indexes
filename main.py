from pathlib import Path
import shutil
import io
import chardet
import os
import codecs
from openpyxl import Workbook

currentPath = os.path.dirname(os.path.abspath(__file__))
outputPath = currentPath + '\\output\\'
shutil.rmtree('output', ignore_errors=True, onerror=None)
Path("output").mkdir(parents=True, exist_ok=True)

index = "236006"

# excelBookHandler = Workbook()
# mainSheet = excelBookHandler.active
# mainSheet.title = index
# cellShift = 2

inputFolder = os.walk("input")
for root, directories, filenames in inputFolder:
    for filename in filenames:
        txt_input_file_path = os.path.join(root, filename)
        if Path(txt_input_file_path).suffix == '.txt':
            bytes = min(32, os.path.getsize(txt_input_file_path))
            raw_file_descriptor = open(txt_input_file_path, 'rb').read(bytes)
            if raw_file_descriptor.startswith(codecs.BOM_UTF8):
                encoding = 'utf-8-sig'
            else:
                result = chardet.detect(raw_file_descriptor)
                encoding = result['encoding']
            input_file_descriptor = io.open(txt_input_file_path, 'r', encoding=encoding)
            input_data = input_file_descriptor.readlines()
            input_file_descriptor.close()

            excel_book_name = str(input_data[0].split(";")[0])
            excelBookHandler = Workbook()
            mainSheet = excelBookHandler.active
            mainSheet.title = excel_book_name

            line_counter = 1
            for address in input_data:
                cellName = "A" + str(line_counter)
                mainSheet[cellName] = address.split(";")[1]
                line_counter += 1

            excel_book_name = outputPath + excel_book_name + ".xlsx"
            excelBookHandler.save(excel_book_name)



