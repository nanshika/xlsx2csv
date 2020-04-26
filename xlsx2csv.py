import pandas as pd
import os
import sys
import click

@click.group()
def cli():
    pass

@cli.command()
@click.argument('input_files', type=click.Path(exists=True), nargs=-1)  # 必須（nargs=-1 にて入力数を任意化）
@click.option('--output_merged_file_name', '-o', default='merged.csv', help='Output merged file name')  # 任意
@click.option('--merged_is_not_requied', '-m', is_flag=True, default=False, help='Set if the merged output file is NOT required')
@click.option('--individuals_is_not_requied', '-i', is_flag=True, default=False, help='Set if the individual output file[s] is/are NOT required')
def xlsx2csv(input_files, output_merged_file_name, merged_is_not_requied, individuals_is_not_requied):
    headers = None
    
    # 入力ファイルが無いか、引数の中に '.xlsx' ファイルが一つもないとエラーを返す
    if len(input_files) == 0 or not any(['.xlsx' in _ for _ in input_files]):
        raise click.BadParameter('Set at least one .xlsx file for input file[s].（対象の .xlsx ファイルを【一つ以上】指定してください。）')
    
    # アウトプットフォルダは入力ファイルの1つ目と同じ。ファイル名のみの指定でも拾えるように。
    output_dir = os.path.dirname(sys.argv[1])
    if len(output_dir) == 0:
        output_dir = os.getcwd()
    
    # 特別な指定がない限り、マージファイルを出力する
    if not merged_is_not_requied:
        f_out0 = open('{}/{}'.format(output_dir, output_merged_file_name), 'w', encoding='utf_8_sig', newline="")
    
    # 入力ファイルごとに開いて .csv ファイルに出力
    for input_file in input_files:
        book = pd.ExcelFile(input_file)
        file_name = os.path.basename(book)
        
        # 特別な指定がない限り、個別のファイルを出力する
        if not individuals_is_not_requied:
            f_out1 = open('{}/{}'.format(output_dir, file_name.replace('.xlsx', '.csv')), 'w', encoding='utf_8_sig', newline="")
        
        for sheet_name in book.sheet_names:
            df = pd.read_excel(input_file, sheet_name=sheet_name, header=headers)
            # 入力ファイル名に .xlsx が含まれないファイルは処理しない。（厳密には '.xlsx' 【を含む】じゃなく【で終わる】とする必要があるが最低限の確認とした。）
            if not '.xlsx' in input_file:
                continue
            
            # 特別な指定がない限り、個別のファイルを出力する
            if not individuals_is_not_requied:
                df.to_csv(f_out1, sep=',', index=False, header=False)
            
            # ファイル名を1列目に挿入
            df.insert(0, 'file_name', file_name)
            
            # シート名を2列目に挿入
            df.insert(1, 'sheet_name', sheet_name)
            
            # 特別な指定がない限り、マージファイルを出力する
            if not merged_is_not_requied:
                df.to_csv(f_out0, sep=',', index=False, header=False)

def main():
    xlsx2csv()

if __name__ == '__main__':
    main()