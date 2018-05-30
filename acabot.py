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

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', SCOPES)


SPREADSHEET = os.environ.get('SPREADSHEET_ID')
CHANNEL_ID = os.environ.get('CHANNEL_ID')

app = Flask(__name__)
app.config['PROPAGATE_EXCEPTIONS'] = True

SLACK_CLIENT = SlackClient(os.environ.get('SLACK_OAUTH_TOKEN'))


def is_valid(request):
    return request.form['token'] == os.environ.get('SLACK_VERIFICATION_TOKEN')


@app.route('/attendance', methods=['POST'])
def attendance():
    client = gspread.authorize(creds)
    logging.info('attendance request: %s', request.form)
    if not is_valid(request):
        print('invalid token ' + request.form['token'])
        abort(400)

    sheet = client.open_by_key(SPREADSHEET).sheet1
    if not request.form['text'].strip():
        return _get_user_attendances(request.form['user_name'], sheet)

    pieces = request.form['text'].split()

    if pieces[0] == 'help':
        return _get_help_text()

    try:
        datetime.datetime.strptime(pieces[0], '%Y-%m-%d')
    except ValueError:
        return jsonify(text='%s doesn\'t look like a date' % pieces[0])

    try:
        date_cell = sheet.find(pieces[0])
        if not date_cell:
            return jsonify(
                text='%s not found in the spreadsheet :(' % pieces[0])
    except gspread.exceptions.CellNotFound:
        return jsonify(text='%s not found in the spreadsheet :(' % pieces[0])

    if len(pieces) == 1:
        return _get_attendance(date_cell, sheet)

    return _record_attendance(date_cell, sheet, pieces)


def _get_user_attendances(username, sheet):
    user_cells = sheet.findall(re.compile('^' + username))
    if not user_cells:
        return jsonify(text='No upcoming attendance records found.')

    lines = []
    for cell in user_cells:
        rowdatestring = sheet.cell(cell.row, 1).value
        rowdate = datetime.datetime.strptime(rowdatestring, '%Y-%m-%d').date()
        if rowdate < datetime.datetime.today().date():
            continue

        pieces = cell.value.split(' | ')
        if len(pieces) > 2 and ' '.join(pieces[2:]):
            text = '*%s:* %s "%s"' % (rowdatestring, pieces[1],
                                      ' '.join(pieces[2:]))
        else:
            text = '*%s:* %s' % (rowdatestring, pieces[1])

        lines.append(text)

    if not lines:
        return jsonify(text='No upcoming attendance records found.')

    return jsonify(text='\n'.join(lines))


def _get_attendance(date_cell, sheet):
    yes = []
    no = []
    maybe = []
    users_found = []
    for cell in sheet.range(date_cell.row, 2, date_cell.row, sheet.col_count):
        if cell.value:
            pieces = cell.value.split(' | ')
            users_found.append(pieces[0])
            if pieces[1] == 'yes':
                yes.append(pieces[0])
            elif pieces[1] == 'maybe':
                maybe.append((pieces[0], ' '.join(pieces[2:])))
            elif pieces[1] == 'no':
                no.append((pieces[0], ' '.join(pieces[2:])))

    no_response = [u for u in _get_active_users() if u not in users_found]
    all_yeses = yes + ['%s (no response)' % u for u in no_response]

    text = '\n'.join([
        '*Yes:*', ', '.join(all_yeses) or '_no entries found_', '', '*Maybe:*',
        ', '.join(['%s "%s"' % (u[0], u[1])
                   for u in maybe]) or '_no entries found_', '', '*No:*',
        ', '.join(['%s "%s"' % (u[0], u[1])
                   for u in no]) or '_no entries found_'
    ])

    return jsonify(text=text)


def _get_active_users():
    channel_info = SLACK_CLIENT.api_call('channels.info', channel=CHANNEL_ID)
    active_members = channel_info['channel']['members']
    all_members = SLACK_CLIENT.api_call('users.list')
    return [
        u['profile']['display_name'] for u in all_members['members']
        if u['id'] in active_members
    ]


def _record_attendance(date_cell, sheet, pieces):
    if pieces[1] not in ['yes', 'no', 'maybe']:
        return jsonify(
            text='2nd argument must be one of "yes", "no", or "maybe"')

    if pieces[1] != 'yes' and len(pieces) < 3:
        return jsonify(text='Reason is required')

    sheet.update_cell(date_cell.row,
                      _get_user_column(sheet, date_cell.row,
                                       request.form['user_name']), ' | '.join([
                                           request.form['user_name'],
                                           pieces[1], ' '.join(pieces[2:])
                                       ]))
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
    text = """
Request all of your upcoming attendances: */attendance*

See attendance for a particular date: */attendance {date}*
    e.g. */attendance 2018-08-01*

Record your attendance: */attendance {date} {yes|no|maybe} {reason}*
    e.g. */attendance 2018-08-01 no on vacation!*
    (reason is only required for "no" and "maybe")
    """
    return jsonify(text=text)
