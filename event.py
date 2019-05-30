# encoding:utf-8
from __future__ import print_function, unicode_literals

import sys
import os
import sqlite3
import random
import unicodedata as ud

import workflow
from workflow import Workflow
from workflow import Workflow3
from consts import *
from util import Util

GITHUB_SLUG = 'xeric/alfred-outlook'
UPDATE_FREQUENCY = 7

log = None

APPLE_SCRIPT = """

set theStartDate to current date
set hours of theStartDate to 0
set minutes of theStartDate to 0
set seconds of theStartDate to 0
set theEndDate to theStartDate + (2 * days) - 1

tell application "Microsoft Outlook"
	set allEvents to every calendar event where its start time is greater than or equal to theStartDate and end time is less than or equal to theEndDate
	
	repeat with aEvent in allEvents
		set eventProps to properties of aEvent
		log eventProps
	end repeat
	
end tell

"""

SELECT_STR = """Select PathToDataFile, Contact_DisplayName, Record_ExchangeOrEasId
        from Contacts 
        where Record_ExchangeOrEasId in (%s)
        """


def main(wf):
    query = wf.decode(sys.argv[1])

    handle(wf, query)

    log.info('searching contact with keyword')


def handle(wf, query):
    if len(query) < 1 or (not str(ud.name(query[0])).startswith("CJK UNIFIED") and len(query) < 2):
        wf.add_item(title='Type more characters to search...',
                    subtitle='too less characters will lead huge irrelevant results',
                    arg='',
                    uid=str(random.random()),
                    valid=False
                    )
    else:
        # run a script to get exchange id
        exchangeIdAndEmail = workflow.util.run_applescript(APPLE_SCRIPT % (query, query,))
        log.info(exchangeIdAndEmail)
        idEmnailList = exchangeIdAndEmail.split(",")

        contacts = None

        log.info("found contact total number: " + str(len(idEmnailList) / 2))

        homePath = os.environ['HOME']

        profile = wf.stored_data(KEY_PROFILE) or OUTLOOK_DEFAULT_PROFILE

        # outlookData = homePath + '/outlook/'
        outlookData = homePath + OUTLOOK_DATA_PARENT + profile + OUTLOOK_DATA_FOLDER
        log.info(outlookData)

        if not Util.validateProfile(outlookData):
            wf.add_item(title='Profile: ' + profile + ' is not valid...',
                        subtitle='please use olkc profile to switch profile',
                        arg='olkc profile ',
                        uid='err' + str(random.random()),
                        valid=False)

        elif len(contacts) == 0:
            wf.add_item(title='No matched contact found...',
                        subtitle='please check the keyword and try again',
                        uid='err' + str(random.random()),
                        valid=False)
        else:
            con = sqlite3.connect(outlookData + OUTLOOK_SQLITE_FILE)
            cur = con.cursor()

            dynamicVarsQM = ""
            dynamicVars = []
            for x in contacts:
                if x[0]:
                    dynamicVarsQM = dynamicVarsQM + "?,"
                    dynamicVars.append(x[0])
            dynamicVarsQM = dynamicVarsQM[:-1]

            res = cur.execute(SELECT_STR % dynamicVarsQM, tuple(dynamicVars))

            resultCount = cur.rowcount
            log.info("got " + str(resultCount) + " results found")

            if resultCount:
                for row in cur:
                    relativePath = row[0]
                    dispName = row[1]
                    excid = row[2] or ""
                    fillContacts(contacts, relativePath, dispName, excid)

            cur.close()

            # id, email, name, relative path
            for i, contact in enumerate(contacts):
                if i >= 20:
                    break

                log.info(contact)
                path = outlookData + contact[3] if contact[3] else None

                email = contact[1]
                it = wf.add_item(title=wf.decode(contact[2] or ""),
                                 subtitle=(email if email and email.strip() else "No Email Found"),
                                 valid=True,
                                 uid=str(contact[0]),
                                 arg=path or email,
                                 type='file' if path else "")
                if not Util.isAlfredV2(wf):
                    mod = it.add_modifier(key='ctrl',
                                          subtitle="Compose a mail to " + email,
                                          arg=email,
                                          valid=True)

    wf.send_feedback()


if __name__ == '__main__':
    wf = Workflow(update_settings={
        'github_slug': GITHUB_SLUG,
        'frequency': UPDATE_FREQUENCY
    })

    if not Util.isAlfredV2(wf):
        wf = Workflow3(update_settings={
            'github_slug': GITHUB_SLUG,
            'frequency': UPDATE_FREQUENCY
        })

    log = wf.logger

    if wf.update_available:
        wf.start_update()

    sys.exit(wf.run(main))
