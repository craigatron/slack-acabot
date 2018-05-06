import logging
import os
import re
import time

import gspread

from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from flask import abort, Flask, jsonify, request
from httplib2 import Http
from slackclient import SlackClient
import datetime

SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', SCOPES)

client = gspread.authorize(creds)

SPREADSHEET = '1Xkkjg7WSrZE7nHRT5OxYl4hpamvJa8D3GvBMLo4wceI'

app = Flask(__name__)
app.config['PROPAGATE_EXCEPTIONS'] = True

slack_client = SlackClient(os.environ.get('SLACK_OAUTH_TOKEN'))

def is_valid(request):
    return request.get_json()['token'] == os.environ.get('SLACK_VERIFICATION_TOKEN')

@app.route('/attendance', methods=['POST'])
def attendance():
    logging.warning('attendance request: %s', request.get_json())
    if not is_valid(request):
        print('invalid token ' + request.get_json()['token'])
        abort(400)

    pieces = request.get_json()['text'].split()

    if pieces[0] == 'help':
        return _get_help_text()

    try:
        datetime.datetime.strptime(pieces[0], '%Y-%m-%d')
    except ValueError:
        return jsonify(text='%s doesn\'t look like a date' % pieces[0])

    sheet = client.open_by_key(SPREADSHEET).sheet1
    date_cell = sheet.find(pieces[0])
    if not date_cell:
        return jsonify(text='%s not found in the spreadsheet :(' % pieces[0])

    if len(pieces) == 1:
        return _get_attendance(date_cell, sheet)

    return _record_attendance(date_cell, sheet, pieces)

def _get_attendance(date_cell, sheet):
    row_values = sheet.row_values(date_cell.row)
    return jsonify(text=' '.join(row_values))


def _record_attendance(date_cell, sheet, pieces):
    if pieces[1] not in ['yes', 'no', 'maybe']:
        return jsonify(text='2nd argument must be one of "yes", "no", or "maybe"')

    if len(pieces) < 3:
        return jsonify(text='Reason is required')

    sheet.update_cell(
            date_cell.row,
            _get_user_column(sheet, date_cell.row, request.get_json()['user_name']),
            ' | '.join([request.get_json()['user_name'], pieces[1], ' '.join(pieces[2:])]))
    return jsonify(text='Recorded!')


def _get_user_column(sheet, row, username):
    first_empty_col = None
    for cell in sheet.range(row, 2, row, sheet.col_count):
        if not cell.value and first_empty_col is None:
            first_empty_col = cell.col

        pieces = cell.value.split(' | ')
        if pieces[0] == username:
            return cell.col

    if first_empty_col is None:
        raise ValueError('No empty columns in row %d' % row)
    return first_empty_col


def _get_help_text():
    return jsonify(text='here\'s some help')
