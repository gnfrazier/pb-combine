import pandas as pd
import glob
import re


'''
Use to combine Performance Merch forms into Apsvia compatible price event
text-tab file.

Input must be on ECOMM tab with column names of:
'Load Pricing (Y/N)','ItemNo','Description','MSRP','Everyday REG Price','Sale Price','','Color','Callout/Feature/Notes'

cols=['Load Pricing (Y/N)','ItemNo','Everyday REG Price','Sale Price']

Output is .txt text tab delim file for upload

TODO: read load pricing flag
      change import to be dynamic based on working directory
      roll up into .exe file
      gui with input box?


'''


def get_files():
    
    files = glob.glob("C:\scripts\combine\*.xlsx")# get the excel file names
    
    return files

def get_suffix(files):
    file_name = re.split(r'\\',files[0]) # parse the first file name from path
    parse = re.split('_', file_name[-1]) # split on _ 
    name = re.sub('\.xlsx','',parse[-1]) # remove the extension
    suffix = parse[0] + '-' + name # Concat the suffix/file name body
    print(suffix)
    return(suffix)

def create_dataframe(files):
    # files = str(r'c:\merge\pb-combine\160711_PB10_Direct_Superflash.xlsx')
    df = pd.DataFrame()
    colnames=['Load Pricing (Y/N)',
          'ItemNo',
          'Description',
          'MSRP',
          'Everyday REG Price',
          'Sale Price',
          '',
          'Color',
          'Callout/Feature/Notes'] # must match names on Merch Form
    cols=['Load Pricing (Y/N)','ItemNo','Everyday REG Price','Sale Price'] # names columns to be imported into dataframe
    dtype= object
    try:
        for f in files:
            data = pd.read_excel(f,
                             sheetname='Ecomm',
                             skiprows=3,
                             usecols=cols,
                             na_values=['',' ',0,0.00,'0','0.00','00.00']
                                 )
            df = df.append(data,ignore_index=True)
    except (ValueError, xlrd.biffh.XLRDError, NameError):
        print("One of the files does not conform to the import template")
    except:
        print("Unexpected error:", sys.exc_info()[0])
        raise    
    return df

def groom_data(df, suffix):
    df.dropna(inplace=True) #strips out blank or bad rows
    df['set']=''    #adds set column (no value)
    df['PageNo']=1  #adds pageno column (default 1)
    df['Suffix']=suffix #adds suffix column
    '''df.rename(index=str, columns = {'Everyday REG Price':'Event Regular Price',
                                   'Sale Price':'Event Sale Price'}) # rename columns for upload file
    '''
    # print(df.head(10))
    return df

def get_file_name(suffix):
    file_name = suffix + ".txt"

    return file_name
                       

def write_txt(all_data, file_name, path_name):
    full_name = path_name + file_name + ".txt" # concat file name from path suffix and extension
    out_columns=['PageNo','Set','ItemNo','Suffix','Everyday REG Price','Sale Price'] # sets the output order
    head_columns=['PageNo','Set','ItemNo','Suffix','Event Regular Price','Event Sale Price']
    all_data.to_csv(path_or_buf = full_name,
                    sep='\t',
                    columns = out_columns,
                    header = head_columns,
                    index=False,
                    float_format="%.2f",
                    dtype={'ItemNo': object,
                           'Everyday REG Price': float,
                           'Event Sale Price': float}) #writes file \t denotes txt tab

    return(full_name)

def write_xlsx(all_data, file_name, path_name):
    full_name = path_name + file_name + ".xlxs" # concat file name from path suffix and extension
    out_columns=['PageNo',
                 'Set',
                 'ItemNo',
                 'Suffix',
                 'Event Regular Price',
                 'Event Sale Price'] # sets the output order
    all_data.to_csv(path_or_buf = full_name,
                    sep='\t',
                    columns = out_columns,
                    header = out_columns,
                    index=False,
                    float_format="%.2f",
                    dtype={'ItemNo': object,
                           'Event Regular Price': float,
                           'Event Sale Price': float}) #writes file \t denotes txt tab

    return(full_name)
            
def main():
    path_name=''
    files = get_files()
    suffix = get_suffix(files)
    file_name = get_file_name(suffix)
    df = create_dataframe(files)
    pfile = groom_data(df, suffix)
    txt_name = write_txt(pfile, suffix, path_name)
    # csv_to_txt(csv_name, suffix, path_name)
    print(txt_name)



if __name__ == '__main__':
    main()
