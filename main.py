# encoding:utf-8
from __future__ import print_function, unicode_literals

import sys
import os
import sqlite3
import re
import random
import unicodedata as ud
from datetime import date, timedelta

import workflow
from workflow import Workflow
from workflow import Workflow3
from consts import *

from util import Util

GITHUB_SLUG = 'xeric/alfred-outlook'
UPDATE_FREQUENCY = 7

log = None

SELECT_STR = """Select Message_NormalizedSubject, Message_SenderList, Message_Preview, PathToDataFile, Message_TimeSent, Message_HasAttachment, Record_RecordID
        from Mail 
        where %s 
        ORDER BY Message_TimeSent DESC 
        LIMIT ? OFFSET ?
        """
FOLDER_COND = """ AND Record_FolderID = ? """
FILTER_COND = """ AND Message_NormalizedSubject NOT LIKE '%s' """

FILTER_COND_FMTED = FILTER_COND % ('')

ATTACH_APPLE_SCRIPT = """
tell application "Microsoft Outlook"
	set msg to get message id %s
	set attas to attachments of msg
	return length of attas
end tell
"""


def main(wf):
    query = wf.decode(sys.argv[1])

    handle(wf, query)

    log.info('searching mail with keyword')


def handle(wf, query):
    # log.info("The query " + query + " is " + str(ud.name(query[0])))
    if len(query) < 2 or (not str(ud.name(query[0])).startswith("CJK UNIFIED") and len(query) < 3):
        wf.add_item(title='Type more characters to search...',
                    subtitle='too less characters will lead huge irrelevant results',
                    arg='',
                    uid=str(random.random()),
                    valid=False
                    )
    else:
        homePath = os.environ['HOME']

        savedFilter = wf.stored_data(KEY_FILTER)
        FILTER_COND_FMTED = FILTER_COND % (savedFilter)

        storedProfle = wf.stored_data(KEY_PROFILE)
        # set default profile to improve user experience
        if storedProfle is None:
            log.info("stored profile is empty, try to get new one!")
            Util.configureDefaultProfile(wf)
        else:
            log.info("configured profile is" + storedProfle)

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
        else :
            # processing query
            page = 0
            if Util.isAlfredV2(wf):
                m = re.search(r'\|(\d+)$', query)
                page = 0 if m is None else int(m.group(1))
            else:
                page = os.getenv('page')
                page = 0 if page is None else int(page)
            if page:
                query = query.replace('|' + str(page), '')
            log.info("query string is: " + query)
            log.info("query page is: " + str(page))

            searchType = 'All'

            originalQuery = query
            if query.startswith('from:'):
                searchType = 'From'
                query = query.replace('from:', '')
            elif query.startswith('title:'):
                searchType = 'Title'
                query = query.replace('title:', '')
            elif query.startswith('recent:'):
                searchType = 'Recent'
                query = query.replace('recent:', '')

            if query is None or query == '':
                wf.add_item(title='Type keywords to search mail ' + searchType + '...',
                            subtitle='too less characters will lead huge irrelevant results',
                            arg='',
                            uid=str(random.random()),
                            valid=True)
            else:
                keywords = query.split(' ')

                configuredPageSize = wf.stored_data('pagesize')
                calculatedPageSize = (int(configuredPageSize) if configuredPageSize else PAGE_SIZE)
                offset = int(page) * calculatedPageSize
                configuredFolder = wf.stored_data('folder')
                folder = (int(configuredFolder) if configuredFolder else 0)

                # start = time()
                con = sqlite3.connect(outlookData + OUTLOOK_SQLITE_FILE)
                cur = con.cursor()

                searchMethod = getattr(sys.modules[__name__], 'query' + searchType)
                searchMethod(cur, keywords, offset, calculatedPageSize + 1, folder)

                resultCount = cur.rowcount
                log.info("got " + str(resultCount) + " results found")

                if resultCount:
                    count = 0

                    for row in cur:
                        count += 1
                        if calculatedPageSize + 1 > count:
                            log.info(row[0])
                            path = outlookData + row[3]
                            if row[2]:
                                content = wf.decode(row[2] or "")
                                content = re.sub('[\r\n]+', ' ', content)
                            else:
                                content = "no content preview"
                            icon = 'attachment.png' if row[5] == 1 else 'mail.png'
                            it = wf.add_item(icon=icon,
                                        title=wf.decode(row[0] or ""),
                                        subtitle=wf.decode('[' + wf.decode(row[1] or "") + '] ' + wf.decode(content or "")),
                                        valid=True,
                                        uid=str(row[4]),
                                        arg=path,
                                        type='file')
                            # download attachment
                            if row[5] == 1 and not Util.isAlfredV2(wf):
                                it.add_modifier(key='ctrl',
                                            subtitle="Click to Download Attachments...",
                                            arg='download_' + str(row[6]),
                                            valid=True)
                    page += 1
                    if count > calculatedPageSize:
                        queryByVersion = originalQuery if not Util.isAlfredV2(wf) else originalQuery + '|' + str(page)
                        it = wf.add_item(title='More Results Available...',
                                    subtitle='click to retrieve next ' + str(calculatedPageSize) + ' results',
                                    arg=queryByVersion,
                                    uid='z' + str(random.random()),
                                    valid=True)
                        if not Util.isAlfredV2(wf):
                            it.setvar('page', page)
                            subtitle = ('no previous page', 'click to retrieve previous ' + str(calculatedPageSize) + ' results')[page > 1]
                            previousPage = 0 if page - 2 < 0 else page - 2
                            mod = it.add_modifier(key='ctrl',
                                            subtitle=subtitle,
                                            arg=queryByVersion,
                                            valid=True)
                            if page > 1:
                                mod.setvar('page', previousPage)
                    else:
                        if page > 1:
                            previousPage = 0 if page - 2 < 0 else page - 2
                            queryByVersion = originalQuery if not Util.isAlfredV2(wf) else originalQuery + '|' + str(previousPage)
                            subtitle = 'click to retrieve previous ' + str(calculatedPageSize) + ' results'
                            it = wf.add_item(title='No More Results',
                                        subtitle=subtitle,
                                        arg=queryByVersion,
                                        uid='z' + str(random.random()),
                                        valid=True)
                            if not Util.isAlfredV2(wf):
                                it.setvar('page', previousPage)
                                mod = it.add_modifier(key='ctrl',
                                                      subtitle=subtitle,
                                                      arg=queryByVersion,
                                                      valid=True)
                                mod.setvar('page', previousPage)

                cur.close()
    wf.send_feedback()


def queryFrom(cur, keywords, offset, pageSize, folder):
    if len(keywords) is None:
        return
    log.info("query by sender")
    log.info(keywords)

    senderConditions = None
    senderVars = []

    for kw in keywords:
        senderVars.append('%' + kw + '%')
        if senderConditions is None:
            senderConditions = '(Message_SenderList LIKE ? '
        else:
            senderConditions += 'AND Message_SenderList LIKE ? '

    senderConditions += ') '

    variables = tuple(senderVars)

    if folder > 0:
        senderConditions += FOLDER_COND
        variables += (folder,)

    senderConditions = "(" + senderConditions
    senderConditions += ")" + FILTER_COND_FMTED

    log.info(SELECT_STR % (senderConditions))
    log.info(variables)

    res = cur.execute(SELECT_STR % (senderConditions), variables + (pageSize, offset,))


def queryTitle(cur, keywords, offset, pageSize, folder):
    if len(keywords) is None:
        return
    log.info("query by subject")
    log.info(keywords)

    titleConditions = None
    titleVars = []

    for kw in keywords:
        titleVars.append('%' + kw + '%')
        if titleConditions is None:
            titleConditions = '(Message_NormalizedSubject LIKE ? '
        else:
            titleConditions += 'AND Message_NormalizedSubject LIKE ? '

    titleConditions += ') '

    variables = tuple(titleVars)

    if folder > 0:
        titleConditions += FOLDER_COND
        variables += (folder,)

    titleConditions = "(" +  titleConditions
    titleConditions += ")" + FILTER_COND_FMTED

    log.info(SELECT_STR % (titleConditions))
    log.info(variables)

    res = cur.execute(SELECT_STR % (titleConditions), variables + (pageSize, offset,))


def queryRecent(cur, keywords, offset, pageSize, folder):
    top = 10

    if len(keywords) is not None:
        if (keywords[0].isnumeric()):
            top = int(keywords[0])
        elif (keywords[0] == 'today'):
            top = -1000

    queryAll(cur, keywords[1:], offset, pageSize, folder, top)

def queryAll(cur, keywords, offset, pageSize, folder, top = -1):
    if len(keywords) == 0 and top == -1:
        return
    log.info("query by subject, content and sender")
    log.info(keywords)

    titleConditions = None
    senderConditions = None
    contentConditions = None
    titleVars = []
    senderVars = []
    contentVars = []
    conditions = " 1 = 1 " # a trick for reduce add additional "AND" in Query

    for kw in keywords:
        titleVars.append('%' + kw + '%')
        senderVars.append('%' + kw + '%')
        contentVars.append('%' + kw + '%')
        if titleConditions is None:
            titleConditions = '(Message_NormalizedSubject LIKE ? '
        else:
            titleConditions += 'AND Message_NormalizedSubject LIKE ? '
        if senderConditions is None:
            senderConditions = 'OR (Message_SenderList LIKE ? '
        else:
            senderConditions += 'AND Message_SenderList LIKE ? '
        if contentConditions is None:
            contentConditions = 'OR (Message_Preview LIKE ? '
        else:
            contentConditions += 'AND Message_Preview LIKE ? '

    if keywords:
        titleConditions += ') '
        senderConditions += ') '
        contentConditions += ') '

        conditions = titleConditions + senderConditions + contentConditions

    variables = tuple(titleVars) + tuple(senderVars) + tuple(contentVars)
    if folder > 0:
        conditions = '(' + conditions + ')' + FOLDER_COND
        variables += (folder,)

    conditions = "(" +  conditions
    conditions += ")" + FILTER_COND_FMTED

    # calculate offset by top value
    # log.info("top: %d, pageSize: %d, offset: %d" % (top, pageSize, offset))
    if top > 0:
        if top > (pageSize - 1):
            if top < (pageSize - 1) + offset:
                pageSize = top - offset
        else:
            pageSize = top

        #append read flag
        if conditions is None:
            conditions = "Message_ReadFlag = 0 "
        else:
            conditions += " AND Message_ReadFlag = 0 "
    elif top == -1000:
        variables += (int(date.today().strftime('%s')),)
        if conditions is None:
            conditions = "Message_TimeReceived > ? AND Message_ReadFlag = 0 "
        else:
            conditions += " AND Message_TimeReceived > ? AND Message_ReadFlag = 0 "


    log.info(SELECT_STR % (conditions))
    log.info(variables + (pageSize, offset,))

    cur.execute(SELECT_STR % (conditions), variables + (pageSize, offset,))


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
