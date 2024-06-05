import json
from flask import Flask, request
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime
import os
# Khởi tạo ứng dụng Flask
app = Flask(__name__)


# Sheet info
credentials = service_account.Credentials.from_service_account_info({
    'type': os.environ['SERVICE_ACCOUNT_TYPE'],
    'project_id': os.environ['SERVICE_ACCOUNT_PROJECT_ID'],
    'private_key_id': os.environ['SERVICE_ACCOUNT_PRIVATE_KEY_ID'],
    'private_key': os.environ['SERVICE_ACCOUNT_PRIVATE_KEY'],
    'client_email': os.environ['SERVICE_ACCOUNT_CLIENT_EMAIL'],
    'client_id': os.environ['SERVICE_ACCOUNT_CLIENT_ID'],
    'auth_uri': os.environ['SERVICE_ACCOUNT_AUTH_URI'],
    'token_uri': os.environ['SERVICE_ACCOUNT_TOKEN_URI'],
    'auth_provider_x509_cert_url': os.environ['SERVICE_ACCOUNT_AUTH_PROVIDER_X509_CERT_URL'],
    'scopes': ['https://www.googleapis.com/auth/spreadsheets']
})
service = build('sheets', 'v4', credentials=credentials)
spreadsheet_id = '1VTi3teBmhHqiJEuv4np7ULxGvP8sPmletdGN-clHNGc'
sheet_name = 'sheet1'


def get_datasheet_raw():
    datasheet_raw = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=sheet_name
    ).execute()
    return datasheet_raw


def update_sheet(range, data):
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range,
        valueInputOption='RAW',
        body=data
    ).execute()
# Định nghĩa route và hàm xử lý


@app.route('/')
def hello_world():
    return 'Hello, World!'


@app.route('/add', methods=['POST'])
def add_post():
    # Xử lý dữ liệu đầu vào
    data_raw = request.get_json()
    date_raw = data_raw['date']
    format_date = datetime.strptime(
        date_raw, '%Y-%m-%dT%H:%M:%S.%fZ').strftime('%d-%m-%Y')

    # Lấy data sheet hiện tại
    datasheet_raw = get_datasheet_raw()
    # Xử lý data sheet hiện tại
    values = datasheet_raw.get('values', [])

    lastrow = len(values)
    lastcol = chr(64+max(len(values[0]), len(values[1])))
    next_index = lastrow
    index_of_date = None
    for index, row in enumerate(values):
        while len(row) < len(values[1]):
            row.append('')
        if row[0] == format_date:
            index_of_date = index
    if index_of_date is not None:
        for j in range(index_of_date+1, lastrow):
            if values[j][0] != '':
                next_index = j

    # Tạo dữ liệu thêm mới
    value_raw = ['']*len(values[1])
    for i in range(len(value_raw)):
        value_raw[i] = ''
    # Check date
    if index_of_date is None:
        value_raw[0] = format_date
    else:
        value_raw[0] = ''
    if data_raw['typeEx'] == 'normal':
        value_raw[1] = data_raw['nameEx']
        value_raw[3] = data_raw['billEx']
    if data_raw['typeEx'] == 'gas':
        value_raw[4] = data_raw['billEx']
    # Thêm dữ liệu vào sheet
    new_range = 'sheet1!A3:'+lastcol+str(lastrow+1)

    if index_of_date is None:
        values.append(value_raw)
    else:
        if index_of_date + 1 == lastrow:
            values.append(value_raw)
        else:
            values.insert(next_index, value_raw)
    new_values = values[2:]
    new_datasheet = datasheet_raw
    new_datasheet['values'] = new_values
    new_datasheet['range'] = new_range

    update_sheet(new_range,new_datasheet)
    
    return f'Hello, {index_of_date, next_index, lastrow}!'


@app.route('/addrevenue', methods=['POST'])
def add_cashinfo():
    data_raw = request.get_json()
    date_raw = data_raw['date']
    format_date = datetime.strptime(
        date_raw, '%Y-%m-%dT%H:%M:%S.%fZ').strftime('%d-%m-%Y')

    # Lấy data sheet hiện tại

    datasheet_raw = get_datasheet_raw()
    # Xử lý data sheet hiện tại
    values = datasheet_raw.get('values', [])
    # values = valuesx
    lastrow = len(values)
    lastcol = chr(64+max(len(values[0]), len(values[1])))
    
    index_of_date = None
    for index, row in enumerate(values):
        while len(row) < len(values[1]):
            row.append('')
        if row[0] == format_date:
            index_of_date = index

    # Tạo dữ liệu thêm mới
    if index_of_date is not None:
        row_for_date = values[index_of_date]
    else:
        row_for_date = ['']*len(values[1])
        row_for_date[0] = format_date
    cash = float(data_raw['cash'])
    deposit = float(data_raw['deposit'])
    numorder = float(data_raw['order'])
    bonus = float(data_raw['bonus'])
    current_asset = cash+deposit
    net_asset = current_asset + numorder*13.5 + bonus

    row_for_date[5] = cash
    row_for_date[6] = deposit
    row_for_date[7] = numorder
    row_for_date[8] = bonus
    row_for_date[13] = current_asset
    row_for_date[14] = net_asset
    # Thêm dữ liệu vào sheet
    if index_of_date is None:
        values.append(row_for_date)
        new_range = 'sheet1!A3:'+lastcol+str(lastrow+1)
    else:
        values[index_of_date] = row_for_date
        new_range = 'sheet1!A3:'+lastcol+str(lastrow)
    new_values = values[2:]
    new_datasheet = datasheet_raw
    new_datasheet['values'] = new_values
    new_datasheet['range'] = new_range
    # run
    update_sheet(new_range,new_datasheet)

    return f'Succes!'
# Chạy ứng dụng
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
