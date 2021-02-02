import json
import random
import re
import statistics
import time

import google.oauth2.service_account
import googleapiclient.discovery


class Feedback:
    url = 'https://sheets.googleapis.com/$discovery/rest?version=v4'

    def __init__(self, master_spreadsheet_id,
                 service_account_file='service-account.json'):
        self.master_spreadsheet_id = master_spreadsheet_id
        print('reading master spreadsheet https://docs.google.com/spreadsheets/d/{}'.format(master_spreadsheet_id))

        credentials = google.oauth2.service_account.Credentials.from_service_account_file(service_account_file)
        self.sheets_api = googleapiclient.discovery.build('sheets', 'v4', credentials=credentials).spreadsheets()
        self.drive_api = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)

        with open('feedback.json') as f:
            config = json.load(f)
        self.master_users = config['master_users']

        for section in ['input', 'results']:
            config[section]['sheet_id'] = \
                self.get_sheet_id(self.master_spreadsheet_id, config[section]['sheet_name'])
            range_ = '{sheet_name}!{topics_col}1:{topics_col}500'.format(**config[section])
            config[section]['last_row'] = len(self.get_range(self.master_spreadsheet_id, range_))
            for key in list(config[section].keys()):
                if key.endswith('_col'):
                    range_ = '{col}2:{col}{last_row}'.format(col=config[section][key], **config[section])
                    config[section][key.replace('_col', '_range')] = range_
        self.config_input = config['input']
        self.config_results = config['results']

        self.title = self.get_title(self.master_spreadsheet_id).replace('Master', '{} - {}')
        self.employee_names = self.get_range(self.master_spreadsheet_id, 'Names!A2:A100')
        self.colleague_names = self.get_range(self.master_spreadsheet_id, 'Names!D2:D100')

    @staticmethod
    def range_length(range_):
        numbers = '-' + re.sub('[a-zA-Z]', '', range_.split('!')[1])
        return sum(int(v) for v in numbers.split(':')) + 1

    def get_title(self, spreadsheet_id):
        return self.get_properties(spreadsheet_id, 'properties/title')['properties']['title']

    def get_properties(self, spreadsheet_id, fields):
        return self.sheets_api.get(spreadsheetId=spreadsheet_id, fields=fields).execute()

    def get_range(self, spreadsheet_id, range_, complete_rows=False):
        values = self.sheets_api.values().get(spreadsheetId=spreadsheet_id,
                                              range=range_).execute().get('values', [])

        def clean(lst):
            v = None
            if len(lst) > 0:
                v = lst[0]
            return v

        if values and len(values[0]) <= 1:
            values = [clean(row) for row in values]
        if complete_rows:
            missing_rows = self.range_length(range_) - len(values)
            values.extend(missing_rows * [None])
        return values

    def update_range(self, spreadsheet_id, range_, values):
        self.sheets_api.values().update(spreadsheetId=spreadsheet_id,
                                        range=range_,
                                        valueInputOption='RAW',
                                        body={'values': values}).execute()

    def create_spreadsheet(self, title, sleep=15):
        new_sheet = self.drive_api.files().create(
            body={
                'name': title,
                'mimeType': 'application/vnd.google-apps.spreadsheet',
            }
        ).execute()
        spreadsheet_id = new_sheet['id']

        time.sleep(sleep)
        for user in self.master_users:
            body = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': user,
            }
            request = self.drive_api.permissions().create(
                fileId=spreadsheet_id,
                body=body,
                fields='id'
            )
            request.execute()
        return spreadsheet_id

    def copy_sheet(self, spreadsheet_id, sheet_id, destination_spreadsheet_id, destination_sheet_title):
        body = {
            'destination_spreadsheet_id': destination_spreadsheet_id
        }
        response = self.sheets_api.sheets().copyTo(spreadsheetId=spreadsheet_id,
                                                   sheetId=sheet_id,
                                                   body=body).execute()
        self.rename_sheet(destination_spreadsheet_id, response['sheetId'], destination_sheet_title)

    def rename_sheet(self, spreadsheet_id, sheet_id, title):
        request = {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "title": title,
                },
                "fields": "title",
            }
        }
        self.batch_update(spreadsheet_id, request)

    def get_sheet_id(self, spreadsheet_id, title):
        properties = self.sheets_api.get(spreadsheetId=spreadsheet_id, fields="sheets.properties").execute()
        for item in properties['sheets']:
            if item['properties']['title'] == title:
                return item['properties']['sheetId']

    def batch_update(self, spreadsheet_id, request):
        body = {
            "requests": [
                request
            ]
        }
        self.sheets_api.batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    def delete_first_sheet(self, spreadsheet_id):
        request = {
            "deleteSheet": {
                "sheetId": 0
            }
        }
        self.batch_update(spreadsheet_id, request)

    def create_feedback_sheets(self):
        spreadsheets_urls = []
        for employee_name in self.employee_names:
            spreadsheets_urls.append([])

            for template in ['Input', 'Results']:
                spreadsheet_title = self.title.format(employee_name, template)
                print('creating spreadsheet "{}"...'.format(spreadsheet_title))
                dest_spreadsheet_id = self.create_spreadsheet(spreadsheet_title)
                spreadsheet_url = 'https://docs.google.com/spreadsheets/d/{}'.format(dest_spreadsheet_id)
                spreadsheets_urls[-1].append(spreadsheet_url)

                if template == 'Input':
                    for dest_sheet_title in self.employee_names:
                        if dest_sheet_title == employee_name:
                            dest_sheet_title = 'Yourself'
                        self.copy_sheet(self.master_spreadsheet_id, self.config_input['sheet_id'],
                                        dest_spreadsheet_id, dest_sheet_title)
                if template == 'Results':
                    self.copy_sheet(self.master_spreadsheet_id, self.config_results['sheet_id'],
                                    dest_spreadsheet_id, 'Results')
                self.delete_first_sheet(dest_spreadsheet_id)
        self.update_range(self.master_spreadsheet_id, 'Names!B2:C100', spreadsheets_urls)

    @staticmethod
    def mean(arr):
        if len(arr) > 0:
            return statistics.mean(arr)
        else:
            return ''

    @staticmethod
    def stddev(arr):
        if len(arr) > 1:
            return statistics.stdev(arr)
        else:
            return ''

    def evaluate_feedback_sheets(self):
        data = self.get_range(self.master_spreadsheet_id, 'Names!A2:C100')
        print(data)
        sheets = {
            item[0]: {
                'input_sheet_id': item[1].split('/')[-1],
                'result_sheet_id': item[2].split('/')[-1]
            }
            for item in data
        }
        print(sheets)
        for employee_name in self.employee_names:
            print('evaluating for "{}"'.format(employee_name))
            team_ratings, team_comments, own_ratings = [], [], []
            for colleague_name in self.employee_names:
                spreadsheet_id = sheets[colleague_name]['input_sheet_id']
                if colleague_name == employee_name:
                    range_ = 'Yourself!{rating_range}'.format(sheet=employee_name, **self.config_input)
                    own_ratings = [int(v) if v is not None else '' for v in self.get_range(spreadsheet_id, range_)]
                else:
                    range_ = '{sheet}!{rating_range}'.format(sheet=employee_name, **self.config_input)
                    team_ratings.append(self.get_range(spreadsheet_id, range_, complete_rows=True))
                    range_ = '{sheet}!{comment_range}'.format(sheet=employee_name, **self.config_input)
                    team_comments.append(self.get_range(spreadsheet_id, range_, complete_rows=True))
            print(team_ratings)
            team_ratings_lists = [[int(v) for v in tr if v is not None] for tr in zip(*team_ratings)]
            team_ratings_mean = [self.mean(tr) for tr in team_ratings_lists]
            team_ratings_stddv = [self.stddev(tr) for tr in team_ratings_lists]
            team_comments_list = ['; '.join(v for v in random.sample(list(c), len(c)) if v) for c in
                                  zip(*team_comments)]

            spreadsheet_id = sheets[employee_name]['result_sheet_id']

            def update(key, data):
                range_ = 'Results!{}'.format(self.config_results[key])
                self.update_range(spreadsheet_id, range_, [[v] for v in data])

            update('team_rating_mean_range', team_ratings_mean)
            update('team_rating_stddev_range', team_ratings_stddv)
            update('oww_rating_range', own_ratings)
            update('team_comment_range', team_comments_list)


if __name__ == '__main__':
    feedback = Feedback('1pJTw1gYVLp1j-BCfz1qXD3IgMX2UwYO90nTeFZ8FaOE')
    feedback.create_feedback_sheets()
    # feedback.evaluate_feedback_sheets()
