# Import Packages
import pandas as pd
import os
from sqlalchemy import create_engine
from datetime import datetime
import configparser
import logging
import time
import sys

# Get Config Filepath - the only file with .ini in directory.
config_filepath = [filename for filename in os.listdir() if '.ini' in filename][0]
print('Config file found: {}'.format(config_filepath))

# Set Logging and Get Config Info
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s')

file_handler = logging.FileHandler('main.log')
file_handler.setFormatter(formatter)

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(stream_handler)


def extract_config(configFile, item):
    try:
        # loads 'Default' as empty string
        if configFile['Authentication'][item] == 'Default':
            logger.info('Loading item {} as Empty String'.format(item))
            return ''

        # throws exception if no entry for desired item
        elif (configItem := configFile['Authentication'][item]) == '':
            logger.exception("Could not find {} in config file.".format(item))

        # if it goes ok, return the info.
        else:
            return configItem
    except KeyError:
        logger.exception('Error extracting {} from config file.'.format(item))


# Try loading config
config = configparser.ConfigParser()
config.read(config_filepath)

# Try loading config items
database_name = extract_config(config, 'database_name')
user = extract_config(config, 'user')
password = extract_config(config, 'password')
connection = extract_config(config, 'connection')
file_folder = extract_config(config, 'folder_path')

# Move to directed folder.
os.chdir(file_folder)

print('Loaded configuration information successfully. \n')

# Create DbConnection
try:
    engine_path = connection.format(user, password, database_name)
    sqlEngine = create_engine(engine_path, pool_recycle=3600)
    dbConnection = sqlEngine.connect()
except Exception:
    logger.exception(
        'Could not connect to MySQL server using the given credentials. Please ensure that the server is up and running, and the credentials are correct.')
    sys.exit(1)


# Define Data Loading Function

def load_data(filename, ext):
    # Load file based on extension.
    if ext == '.xlsx':
        df = pd.read_excel(filename)
    elif ext == '.xlsb':
        df = pd.read_excel(filename, engine='pyxlsb')
    else:
        df = pd.read_csv(filename, sep=',')

    # Add load date column.
    df['load_date'] = datetime.now()
    df['load_date'] = df['load_date'].astype('datetime64[ns]')

    # Drop index column if exists.
    if 'index' in df.columns:
        df.drop('index', inplace=True, axis=1)
        pass

    # Drop table_name column if exists.
    if 'table_name' in df.columns:
        df.drop('table_name', inplace=True, axis=1)

    # Creates table name and adds it to data.
    table_name = os.path.splitext(filename)[0].replace(' ', '_').lower().replace('.', '_')
    df['table_name'] = table_name.lower()

    return df, table_name


def main():
    # creates list of only 'data' files.
    file_list = [filename for filename in os.listdir() if os.path.splitext(filename)[1] in ['.csv', '.xlsx', '.xlsb']]

    print('List of files to be loaded:')
    for fileIndex, fileName in enumerate(file_list):
        print('[{}] {}'.format(fileIndex, fileName))

    for fileIndex, fileName in enumerate(file_list):
        # file_index = int(input('\nWhat file (index) would you like?\n'))

        # print("\nYou've selected the following file:\n{}".format(file_list[file_index]))

        # fileName = file_list[file_index]

        # gets file extension
        ext = os.path.splitext(fileName)[1]

        df, table_name = load_data(fileName, ext)
        df.to_sql(table_name, dbConnection, if_exists='append')

        del df

    print('Finished successfully.')
    time.sleep(5)


if __name__ == '__main__':
    try:
        main()
    except Exception:
        logger.exception('Main Loop: Unhandled Error as follows.')
