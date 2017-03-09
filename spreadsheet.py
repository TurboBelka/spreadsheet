import argparse
import getopt
import os
import httplib2
import logging

from apiclient import discovery
import json
from googleapiclient.errors import HttpError
from oauth2client import client, tools
from oauth2client.client import HttpAccessTokenRefreshError
from oauth2client.file import Storage
from faker import Faker

# spreadSheetId = 1G6T5Rqxg1g0QqPBrZ0ZnU1tkfj73WEJssxKFTYSNKZ8
# --range_name namedRangeFirst --spread_sheet_id 1G6T5Rqxg1g0QqPBrZ0ZnU1tkfj73WEJssxKFTYSNKZ8

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCOPE = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, 'spreadsheet', 'client_secret.json')
ROWS_COUNT = 150
COLUMNS_COUNT = 50

parser = argparse.ArgumentParser()
parser.add_argument("-r", "--range_name",
					help="enter range where you want load data (Named range or cell address)")
parser.add_argument("-s", "--spread_sheet_id",
					help="enter google spreadsheet ID")


logging.basicConfig(format=u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s '
						   u'[%(asctime)s]  %(message)s',
					level=logging.DEBUG)
fh = logging.FileHandler('logs.txt')
fh.setLevel(logging.DEBUG)


def get_credentials():
	credential_dir = os.path.join(BASE_DIR, 'credentials')
	if not os.path.exists(credential_dir):
		os.makedirs(credential_dir)
	credential_path = os.path.join(credential_dir, 'credentials.json')
	store = Storage(credential_path)
	creds = store.get()
	if not creds or creds.invalid:
		flow = client.flow_from_clientsecrets(
			CLIENT_SECRET_FILE,
			scope=SCOPE)
		flow.params['access_type'] = 'offline'  # offline access
		creds = tools.run_flow(flow, store)
	return creds


def generate_fake_data():
	logging.info('START GENERATE DATA')
	fake = Faker()
	fake_data = [[fake.name() for j in xrange(COLUMNS_COUNT)] for i in
				 xrange(ROWS_COUNT)]
	logging.info('Faked Data: %s' % fake_data)
	return fake_data


if __name__ == "__main__":
	try:
		flags = parser.parse_args()
	except getopt.GetoptError as e:
		logging.error('ARG PARSER ERROR: %s' % e)
		flags = None

	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
	service = discovery.build('sheets', 'v4', http=http,
							  discoveryServiceUrl=discoveryUrl)

	fake_data = generate_fake_data()
	body = {'values': fake_data}
	try:
		result = service.spreadsheets().values().update(
			spreadsheetId=flags.spread_sheet_id, range=flags.range_name,
			valueInputOption='USER_ENTERED', body=body).execute()
		logging.info('HTTP RESULT: %s' % result)
	except HttpError as e:
		logging.error('HTTP ERROR: %s' % e.resp)
		logging.error('ERROR MESSAGE: %s' % json.loads(e.content).get('error').get(
			'message'))
	except HttpAccessTokenRefreshError as e:
		logging.error('HTTP ACCESS TOKEN ERROR: %s' % e)
