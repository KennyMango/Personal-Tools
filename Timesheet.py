import pandas as pd

path = 'files/timesheet.xlsx'

def readTimesheet(filepath):
    # Load the Excel spreadsheet into a pandas DataFrame object
    df = pd.read_excel(filepath)
    return df

def insertData(data, df, row, column):
    df.at[row, column] = data
    return df

def main():
    df = readTimesheet(path)
    dfinserted = insertData('Test', df, 38, 'H')
    dfinserted.to_excel(path, index=False)

if __name__ == "__main__":
    main()
