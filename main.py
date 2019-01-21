#encoding:utf-8
from __future__ import print_function, unicode_literals

import sys
import os
import sqlite3
import re
from time import time
import random

from workflow import Workflow
from workflow import util
from consts import *

log = None

SELECT_STR = """Select Message_NormalizedSubject, Message_SenderList, Message_Preview, PathToDataFile, Message_TimeSent
        from Mail 
        where %s 
        ORDER BY Message_TimeSent DESC 
        LIMIT ? OFFSET ?
        """
FOLDER_COND = """ AND Record_FolderID = ? """

def main(wf):
    query = wf.decode(sys.argv[1])

    handle(wf, query)

    log.info('searching mail with keyword')

def handle(wf, query):
    if (len(query) < 3):
        wf.add_item(title='Type more characters to serach...', 
                    subtitle='too less characters will lead huge irrelevant results', 
                    arg='', 
                    uid=str(random.random()), 
                    valid=True
                    )
    else:
        homePath = os.environ['HOME']

        # outlookData = homePath + '/outlook/'
        outlookData = homePath + r'/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/'
        log.info(outlookData)

        # processing query
        m = re.search(r'\|(\d+)$', query)
        page = 0 if m is None else int(m.group(1))
        if page:
            query = query.replace('|' + str(page), '')
        log.info("query string is: " + query)
        log.info("query page is: " + str(page))

        searchType = 'All'

        if query.startswith('from:'):
            searchType = 'From'
            query = query.replace('from:', '')
        elif query.startswith('title:'):
            searchType = 'Title'
            query = query.replace('title:', '')

        if query is None or query == '':
            wf.add_item(title='Type keywords to serach mail ' + searchType + '...', 
                    subtitle='too less characters will lead huge irrelevant results', 
                    arg='', 
                    uid=str(random.random()), 
                    valid=True
                    )
        else:
            keywords = query.split(' ')

            configuredPageSize = wf.stored_data('pagesize')
            calculatedPageSize = (int(configuredPageSize) if configuredPageSize else PAGE_SIZE)
            offset = int(page) * calculatedPageSize
            configuredFolder = wf.stored_data('folder')
            folder = (int(configuredFolder) if configuredFolder else 0)

            """Read in the data source and add it to the search index database"""
            # start = time()
            con = sqlite3.connect(outlookData + 'Outlook.sqlite')
            count = 0
            cur = con.cursor()

            searchMethod = getattr(sys.modules[__name__], 'query' + searchType)
            searchMethod(cur, keywords, offset, calculatedPageSize, folder)

            if cur.rowcount:
                for row in cur:
                    log.info(row[0])
                    path = outlookData + row[3]
                    if row[2]:
                        content = wf.decode(row[2])
                        content = content.replace('\r\n', " ")
                    else:
                        content = "no content preview"
                    wf.add_item(title=unicode(row[0]), subtitle=unicode('[' + unicode(row[1]) + '] ' + unicode(content)), valid=True, uid=str(row[4]), arg=path, type='file')
                page += 1        
                wf.add_item(title='Next ' + str(calculatedPageSize) + ' results...', subtitle='click to retrieve another ' + str(calculatedPageSize) + ' results', arg=query + '|' + str(page), uid='z' + str(random.random()), valid=True)

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
        variables += (folder, )

    log.info(SELECT_STR % (senderConditions))
    log.info(variables)

    res = cur.execute( SELECT_STR % (senderConditions), variables + (pageSize, offset, ))

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
        variables += (folder, )

    log.info(SELECT_STR % (titleConditions))
    log.info(variables)

    res = cur.execute( SELECT_STR % (titleConditions), variables + (pageSize, offset, ))

def queryAll(cur, keywords, offset, pageSize, folder):
    if len(keywords) is None:
        return
    log.info("query by subject, content and sender")
    log.info(keywords)

    titleConditions = None
    senderConditions = None
    contentConditions = None
    titleVars = []
    senderVars = []
    contentVars = []

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

    titleConditions += ') '
    senderConditions += ') '
    contentConditions += ') '

    conditions = titleConditions + senderConditions + contentConditions

    variables = tuple(titleVars) + tuple(senderVars) + tuple(contentVars)
    if folder > 0:
        conditions = '(' + conditions + ')' + FOLDER_COND
        variables += (folder, )

    log.info(SELECT_STR % (conditions))
    log.info(variables + (pageSize, offset, ))

    res = cur.execute( SELECT_STR % (conditions), 
        variables + (pageSize, offset, ))

if __name__ == '__main__':
    wf = Workflow()
    log = wf.logger
    sys.exit(wf.run(main))