import pandas as pd
import re
import logging
import math
import numpy as np
from numpy import ceil
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# API and spreadsheet settings
API_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_ID = '1Tx0swikem8WWTHEqTrvIB4okJpLEQ9fpLogy4K02JiY'
DATA_RANGE = 'engenharia_de_software!A3:F27'
RESULT_RANGE = 'engenharia_de_software!G4:H27'
CREDENTIALS_FILE = 'tunts-rock-2024-psilva8769.json'

def authenticate_with_google(api_scopes, credentials_path):
    # Authenticate with the Google API using the service account credentials file.
    logging.info('Authenticating with Google Sheets API...')
    credentials = service_account.Credentials.from_service_account_file(credentials_path, scopes=api_scopes)
    return build('sheets', 'v4', credentials=credentials)

# Extract
def fetch_data_from_sheet(sheet_service, sheet_id, range_name):
    # Fetches data from the specified spreadsheet
    logging.info(f'Fetching data from sheet: {range_name}')
    values = sheet_service.spreadsheets().values()
    total_classes = values.get(spreadsheetId=sheet_id, range='engenharia_de_software!A2').execute()
    result = values.get(spreadsheetId=sheet_id, range=range_name).execute()
    data = result.get('values', [])
    return [pd.DataFrame(data[1:], columns=data[0]), total_classes]

# Transform
def process_student_data(df, total_classes_string):
    # Processes student data to determine status and final grade
    logging.info('Processing student data...')
    total_classes_string_list = re.findall(r'\d+', total_classes_string['values'][0][0])
    total_classes_float = [float(number) for number in total_classes_string_list]
    allowed_absences = total_classes_float[0]*0.25
    print(math.ceil(allowed_absences))

   # Clean column names and convert columns to numeric where possible
    df = df.rename(columns=lambda x: x.strip()).apply(lambda x: pd.to_numeric(x, errors='ignore'))

    # Calculate average
    df['Average'] = (df['P1'] + df['P2'] + df['P3']) / 3

    # Initialize 'Status' and 'Final Grade' columns
    df['Status'] = np.nan
    df['Final Grade'] = 0

    # Apply conditions
    df.loc[df['Average'].between(50, 70, inclusive='left'), 'Status'] = 'Exame Final'
    df.loc[df['Average'] < 50, 'Status'] = 'Reprovado por Nota'
    df.loc[df['Faltas'].astype(int) > ceil(allowed_absences), 'Status'] = 'Reprovado por Falta'

    # Calculate 'Final Grade' for 'Exame Final' status
    df.loc[df['Status'] == 'Exame Final', 'Final Grade'] = ((df['Average'] + (70 - df['Average'])) / 2).apply(ceil)

    # For rows that dont fall into 'Exame Final, this makes sure that 'Final Grade' stays 0
    # Kinda of an optinal step but useful if changes are ever needed
    df.loc[df['Status'] != 'Exame Final', 'Final Grade'] = 0

    # For any rows that didn't meet any condition, set a default status
    df['Status'].fillna('Aprovado', inplace=True)

    return df[['Status', 'Final Grade']]

# Load
def update_sheet(sheet_service, sheet_id, range_name, data):
    # Updates the spreadsheet with the processed data
    logging.info('Updating the spreadsheet with new data...')
    body = {'values': data.values.tolist()}
    response = sheet_service.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=range_name, valueInputOption='RAW', body=body).execute()
    logging.info(f"{response.get('updatedCells')} cells updated.")

def main():
    try:
        # Authenticate with Google API
        service = authenticate_with_google(API_SCOPES, CREDENTIALS_FILE)

        # Extract: Fetch data from Google Sheets
        student_data = fetch_data_from_sheet(service, SHEET_ID, DATA_RANGE)
        
        # Transform: Process the data to determine student status and final grade
        processed_data = process_student_data(student_data[0], student_data[1])

        # Load: Update the Google Sheet with the processed data
        update_sheet(service, SHEET_ID, RESULT_RANGE, processed_data)
        
    except HttpError as error:
        logging.error(f'An error occurred: {error}')

if __name__ == '__main__':
    main()
