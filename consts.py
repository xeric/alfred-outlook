#coding:utf-8
PAGE_SIZE = 20
OUTLOOK_DATA_PATH = r'/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/'

KEY_PAGE_SIZE = 'pagesize'
KEY_FOLDER = 'folder'
ALL_KEYS = [KEY_PAGE_SIZE, KEY_FOLDER]
ALL_KEY_DESCS = ['one page result size', 'search folder']
ALL_VALS = ['[number]', '[folder name]']

RULES = [r'pagesize \d+', r'folder \d+']