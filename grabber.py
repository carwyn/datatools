import sys
import argparse

import pandas as pd

def extract_columns(filename, columns=[], header=None, parse_cols=None,
        sheetname=None):
    """Extract the specified data from the spreadsheet."""
    xlsx = pd.ExcelFile(filename)

    df = xlsx.parse(sheetname=sheetname, header=header,
        parse_cols=parse_cols)

    for c in columns:
        if c not in df.columns:
            raise KeyError('Requested column not found: {0}'.format(c))

    return df[columns]

def write_columns(dataframe, output=sys.stdout, sep=','):
    """Write out the specified dataset to a csv file."""
    
    dataframe.to_csv(output, sep=sep, index=False)

def parse_args():

    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Grab columns out of OOXML xlsx files.')
    parser.add_argument('-f', '--filename', type=argparse.FileType('r'),
        help='Input file.', required=True)
    parser.add_argument('-o', '--output', default=sys.stdout,
        type=argparse.FileType('w'), help='Output file.')
    parser.add_argument('columns', nargs='+',
        help='Columns to select.')
    parser.add_argument('-d', '--header', type=int, 
        help='Header rows to skip.')
    parser.add_argument('-p', '--parsecols', type=int,
        help='Limit the number of columns to return.')
    parser.add_argument('-w', '--sheetname', required=True,
        help='Name of worksheet to extract from.')
    parser.add_argument('-s', '--separator', default=',',
        help='Field separator for output file.')

    return parser.parse_args()

def main():
    """Main entry point."""
    args = parse_args()
    df = extract_columns(args.filename,
                            columns=args.columns,
                            header=args.header,
                            parse_cols=args.parsecols,
                            sheetname=args.sheetname)

    write_columns(df, sep=args.separator)

if __name__ == '__main__':
    main()

