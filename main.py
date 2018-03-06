#coding:utf-8
from __future__ import print_function, unicode_literals

import sys
import os
import sqlite3
import re
from time import time
import random

from workflow import Workflow

log = None

PAGE_SIZE = 10

SELECT_STR = """Select Message_NormalizedSubject, Message_SenderList, Message_Preview, PathToDataFile, Message_TimeSent
        from Mail 
        where %s 
        ORDER BY Message_TimeSent DESC 
        LIMIT ? OFFSET ?
        """

def main(wf):
    query = sys.argv[1]
    log.info(query)

    handle(query)

    log.info('searching mail with keyword')

def handle(query):
    if (len(query) < 3):
        return

    homePath = os.environ['HOME']

    # outlookData = homePath + '/outlook/'
    outlookData = homePath + r'/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/'
    log.info(outlookData)

    query = unicode(query, 'utf-8')  

    # processing query
    m = re.search(r'\|(\d+)$', query)
    page = 0 if m is None else int(m.group(1))
    if page:
        query = query.replace('|' + str(page), '')
    log.info("query string is: " + unicode(query))
    log.info("query page is: " + str(page))

    searchType = 'All'

    if query.startswith('from:'):
        searchType = 'From'
        query = query.replace('from:', '')
    elif query.startswith('title:'):
        searchType = 'Title'
        query = query.replace('title:', '')

    if query is None or query == '':
        return

    keywords = query.split(' ')

    configuredPageSize = wf.stored_data('page_size')
    offset = int(page) * (configuredPageSize if configuredPageSize else PAGE_SIZE)

    """Read in the data source and add it to the search index database"""
    # start = time()
    con = sqlite3.connect(outlookData + 'Outlook.sqlite')
    count = 0
    cur = con.cursor()

    searchMethod = getattr(sys.modules[__name__], 'query' + searchType)
    searchMethod(cur, keywords, offset)

    if cur.rowcount: 
        for row in cur:
            count
            log.info(row[0])
            path = outlookData + row[3]
            if row[2]:
                content = row[2].encode('ascii', 'ignore').decode('ascii')
                content = content.replace('\r\n', " ")
            else:
                content = "no content preview"
            wf.add_item(title=unicode(row[0]), subtitle=str('[' + unicode(row[1]) + '] ' + content), valid=True, uid=str(row[4]), arg=path, type='file')
        page += 1        
        wf.add_item(title='Next ' + str(PAGE_SIZE) + ' results...', subtitle='click to retrieve another ' + str(PAGE_SIZE) + ' results', arg=query + '|' + str(page), uid='z' + str(random.random()), valid=True)

    cur.close()
    wf.send_feedback()

def queryFrom(cur, keywords, offset):
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

    log.info(SELECT_STR % (senderConditions))
    log.info(variables)

    res = cur.execute( SELECT_STR % (senderConditions), variables + (PAGE_SIZE, offset, ))

def queryTitle(cur, keywords, offset):
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

    log.info(SELECT_STR % (titleConditions))
    log.info(variables)

    res = cur.execute( SELECT_STR % (titleConditions), variables + (PAGE_SIZE, offset, ))

def queryAll(cur, keywords, offset):
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

    variables = tuple(titleVars) + tuple(senderVars) + tuple(contentVars)

    log.info(SELECT_STR % (titleConditions + senderConditions + contentConditions))
    log.info(variables)

    res = cur.execute( SELECT_STR % (titleConditions + senderConditions + contentConditions), 
        variables + (PAGE_SIZE, offset, ))

if __name__ == '__main__':
    wf = Workflow()
    log = wf.logger
    sys.exit(wf.run(main))