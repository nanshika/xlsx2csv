import pandas as pd
import os
import sys

if len(sys.argv) > 1:
    headers = None
    if len(sys.argv[1]) > 1:
        output_dir = os.path.dirname(sys.argv[1])
    else:
        output_dir = os.getcwd()
    
    f_out = open( output_dir + '/exported.csv', 'w', encoding='utf_8_sig', newline="")
    for i in range(1, len(sys.argv)):
        f_in = sys.argv[i]
        book = pd.ExcelFile(f_in)
        file_name = os.path.basename(book)
        for sheet_name in book.sheet_names:
            df = pd.read_excel(f_in, sheet_name=sheet_name, header=headers)
            df.insert(0, 'file_name', file_name)  # ファイル名を1列目に挿入
            df.insert(1, 'sheet_name', sheet_name)  # シート名を2列目に挿入
            df.to_csv(f_out, sep=',', index=False, header=False)
