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
tell application "Microsoft Outlook"

    set results to {}

    repeat with aContact in contacts
        set itemProps to properties of aContact
        log itemProps

        set emails to email addresses of aContact

        if (display name of aContact contains "%s") or my emailContains(email addresses of aContact, "%s") then
            set exchangeId to exchange id of aContact
            set emails to email addresses of aContact
            if exchangeId is missing value then set exchangeId to ""
            if emails is missing value or emails is {} then
                set end of results to exchangeId
                set end of results to ""
            else
                repeat with email in emails
                    set end of results to exchangeId
                    set end of results to address of email
                end repeat
            end if
            # set end of results to display name of aContact & exchange id of aContact & item 1 of email addresses of aContact
        end if

    end repeat

    return results

end tell

on emailContains(emails, keyword)
    tell application "Microsoft Outlook"
        set check to false
        if emails is not missing value or emails is not {} then
            repeat with email in emails
                if (address of email contains keyword) then
                    set check to true
                    exit repeat
                end if
            end repeat
        end if
        return check
    end tell
end emailContains

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

        contacts = buildContacts(idEmnailList)

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


def buildContacts(idEmail):
    contacts = []
    for i, k in zip(idEmail[0::2], idEmail[1::2]):
        contacts.append([i.strip(), k.strip(), k.strip(), None])

    return contacts


def fillContacts(contacts, path, name, id):
    for i, x in enumerate(contacts):
        if x[0] == id and x[3] is None:
            contacts[i][2] = name
            contacts[i][3] = path


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
