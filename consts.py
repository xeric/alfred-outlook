# coding:utf-8
PAGE_SIZE = 20
# OUTLOOK_DATA_PATH = r'/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/'
OUTLOOK_DATA_PARENT = r'/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/'
OUTLOOK_DEFAULT_PROFILE = r'Main Profile'
OUTLOOK_DATA_FOLDER = r'/Data/'
OUTLOOK_SQLITE_FILE = r'Outlook.sqlite'

KEY_PAGE_SIZE = 'pagesize'
KEY_FOLDER = 'folder'
KEY_PROFILE = 'profile'
ALL_KEYS = [KEY_PAGE_SIZE, KEY_FOLDER, KEY_PROFILE]
ALL_KEY_DESCS = ['one page result size', 'search folder', 'outlook profile']
ALL_VALS = ['[number]', '[folder name]', '[profile name]']

RULES = [r'(pagesize)\s+(\d+)', r'(folder)\s+(\d+)', r'(profile)\s+([\w\s_]+)']