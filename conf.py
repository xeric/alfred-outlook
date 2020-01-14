#coding:utf-8
from __future__ import print_function, unicode_literals

import sys
import os
import re

from workflow import Workflow
from consts import *

log = None


def main(wf):
    query = wf.decode(sys.argv[1])
    log.info(query)

    handle(wf, query)


def handle(wf, query):
    query = query.strip()

    if re.match(r'\S+\s+\S+', query) is None:
        hasOption = False
        for i, key in enumerate(ALL_KEYS):
            if key.startswith(query):
                if query == KEY_FOLDER:
                    hasOption = True
                    prepareFolders()
                elif query == KEY_PROFILE:
                    hasOption = True
                    prepareProfile()
                else:
                    hasOption = True
                    value = wf.stored_data(key)
                    option = key + ' ' + "[" + value + "]" if value else ALL_VALS[i]
                    wf.add_item(
                        title=option, 
                        subtitle="Set Config for " + ALL_KEY_DESCS[i], 
                        arg=key, 
                        uid=option, 
                        valid=True
                        )
        if not hasOption:
            option = 'You can configure: ' + ', '.join(ALL_KEYS)
            wf.add_item(
                title=option, 
                subtitle='', 
                arg=option, 
                uid=option, 
                valid=False
                )
    else:
        query.replace(r'\s+', ' ')
        kvPair = query.split(' ')

        key = kvPair[0]
        val = kvPair[1]

        message = None
        if ALL_KEYS.index(key) >= 0:
            # wf.store_data(key, val)
            # message = 'Configure ' + key + ' to ' + val + 'successfully!'
            arg = key + ' ' + val
            wf.add_item(
                title='Set ' + ALL_KEY_DESCS[ALL_KEYS.index(key)] + ' to ' + val, 
                subtitle='', 
                arg=arg, 
                uid=arg, 
                valid=True
                )
        else:
            message = 'Unrecognize configuration name: ' + key

    wf.send_feedback()

def prepareFolders():
    import sqlite3

    #add default folder
    wf.add_item(
        title='Default(All)', 
        subtitle='Set search folder under Default(All)', 
        valid=True, 
        uid='folder 0', 
        arg='folder 0', 
        )

    homePath = os.environ['HOME']

    profile = wf.stored_data(KEY_PROFILE) or OUTLOOK_DEFAULT_PROFILE

    # outlookData = homePath + '/outlook/'
    outlookData = homePath + OUTLOOK_DATA_PARENT + profile + OUTLOOK_DATA_FOLDER
    log.info(outlookData)

    con = sqlite3.connect(outlookData + 'Outlook.sqlite')
    count = 0
    cur = con.cursor()

    log.info("start querying folders")
    cur.execute( """
                SELECT Record_RecordID, Folder_Name FROM Folders 
                WHERE Record_AccountUID > 0
                AND Folder_Name is not null 
            """)

    if cur.rowcount: 
        for row in cur:
            log.info(row[1])
            wf.add_item(
                title=unicode(row[1]), 
                subtitle='Set search folder under ' + unicode(row[1]), 
                valid=True, 
                uid='folder ' + str(row[0]), 
                arg='folder ' + str(row[0]), 
                )

    cur.close()

def prepareProfile():
    homePath = os.environ['HOME']

    parent = homePath + OUTLOOK_DATA_PARENT
    profiles = os.listdir(parent)
    for profile in profiles:
        if not profile.startswith('.'):
            wf.add_item(
                    title=unicode(profile), 
                    subtitle='Set profile as: ' + unicode(profile), 
                    valid=True, 
                    uid='profile ' + str(profile), 
                    arg='profile ' + str(profile), 
                    )

if __name__ == '__main__':
    wf = Workflow()
    log = wf.logger
    sys.exit(wf.run(main))