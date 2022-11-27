## Import Packages
import pandas as pd
import os
from sqlalchemy import create_engine
import pymysql
import pyxlsb
from datetime import datetime
import configparser
import logging

## Get Config Filepath

file_list = [filename for filename in os.listdir() if '.ini' in filename]
print('List of config files in directory:')
for index, filename in enumerate(file_list):
    print('[{}] {}'.format(index, filename))

file_index = int(input('\nEnter the index of your config file.\n'))

config_filepath = file_list[file_index]

## Set Logging and Get Config Info

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s')

file_handler = logging.FileHandler('main.log')
file_handler.setFormatter(formatter)

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(stream_handler)


def extract_config(config, item):
    try:
        if config['Authentication'][item] == 'Default':
            logger.info('Loading item {} as Empty String'.format(item))
            return ''
        elif (auth_info := config['Authentication'][item]) == '':
            logger.exception("Could not find {} in config file.".format(item))
        else:
            return auth_info
    except Exception as error:
        logger.exception('Error extracting {} from config file.'.format(item))


config = configparser.ConfigParser()
config.read(config_filepath)

database_name = extract_config(config, 'database_name')
user = extract_config(config, 'user')
password = extract_config(config, 'password')
connection = extract_config(config, 'connection')

print('Loaded configuration information successfully. \n')

## Create DbConnection
try:
    engine_path = connection.format(user, password, database_name)
    sqlEngine = create_engine(engine_path, pool_recycle=3600)
    dbConnection = sqlEngine.connect()
except:
    logger.exception(
        'Could not connect to MySQL server using the given credentials. Please ensure that the server is up and running, and the credentials are correct.')


## Define Data Loading Function

def load_data(filename, ext):
    if ext == '.xlsx':
        sheet = input('\nWhat is the sheetname? Hit enter to indicate first sheet.\n')
        if sheet == '':
            df = pd.read_excel(filename)
        else:
            df = pd.read_excel(filename, sheet_name=sheet)
    elif ext == '.xlsb':
        print('\nPlease convert to an excel, unless you are ok with your dates being messed up!')
        sheet = input('\nWhat is the sheetname? Hit enter to indicate first sheet.\n')
        if sheet == '':
            pd.read_excel(filename, engine='pyxlsb')
        else:
            df = pd.read_excel(filename, sheet_name=sheet, engine='pyxlsb')
    else:
        sep = input('\nWhat is the seperator? Hit enter to default to ",".\n')
        if sep == '':
            df = pd.read_csv(filename, sep=',')
        else:
            df = pd.read_csv(filename, sep=sep)

    df['load_date'] = datetime.now()
    df['load_date'] = df['load_date'].astype('datetime64[ns]')

    table_name = input('\nDesired table name? Hit enter to load from file name.\n')

    if 'index' in df.columns:
        df.drop('index', inplace=True, axis = 1)
        pass
    if table_name == '':
        table_name = os.path.splitext(filename)[0].replace(' ', '_').lower().replace('.', '_')

    if input('\nWould you like to add the table name as a column? Hit enter to default to Yes.') == '':
        if 'table_name' in df.columns:
            df.drop('table_name', inplace=True, axis=1)
        df['table_name'] = table_name.lower()

    return df, table_name


## Main While Loop

def main():
    # excludes non-interesting files.
    file_list = [filename for filename in os.listdir() if
                 '.ipynb' not in filename and '.py' not in filename and '.ini' not in filename and '.log' not in filename and '.exe' not in filename]
    print('List of files in directory:')
    for index, filename in enumerate(file_list):
        print('[{}] {}'.format(index, filename))

    file_index = int(input('\nWhat file (index) would you like?\n'))

    print("\nYou've selected the following file:\n{}".format(file_list[file_index]))

    filename = file_list[file_index]

    # gets file extension
    ext = os.path.splitext(filename)[1]

    if ext == '.xlsx' or ext == '.xlsb':
        while True:
            df, table_name = load_data(filename, ext)
            frame = df.to_sql(table_name, dbConnection, if_exists='replace')
            if input(
                    'Would you like to get another sheet from this file? Enter to quit or enter "C" for continue. \n') != 'C':
                break
            else:
                pass
    else:
        if ext != '.csv':
            if input(
                    'WARNING: The file you are loading may not be a delimited text file. If you would like to go ahead, please press Enter. Otherwise, input "Q"') != 'Q':
                df, table_name = load_data(filename, ext)
                frame = df.to_sql(table_name, dbConnection, if_exists='replace')
            else:
                pass
        else:
            df, table_name = load_data(filename, ext)
            frame = df.to_sql(table_name, dbConnection, if_exists='replace')


if __name__ == '__main__':
    while True:
        # noinspection PyBroadException
        try:
            main()
        except Exception:
            logger.exception('Main Loop: Unhandled Error as follows.')
