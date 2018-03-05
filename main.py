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

def main(wf):
    query = sys.argv[1]  

    handle(query)

    log.info('searching mail with keyword')

def handle(query):
    if (len(query) < 3):
        return

    query = unicode(query, 'utf-8')  

    homePath = os.environ['HOME']
    # outlookData = homePath + '/outlook/'
    outlookData = homePath + r'/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/'
    log.info(outlookData)

    """Read in the data source and add it to the search index database"""
    # start = time()
    con = sqlite3.connect(outlookData + 'Outlook.sqlite')
    count = 0
    cur = con.cursor()

    m = re.search(r'\|(\d+)$', query)
    page = 0 if m is None else int(m.group(1))
    if page:
        query = query.replace('|' + str(page), '')
    log.info("query string is: " + unicode(query))
    log.info("query page is: " + str(page))

    offset = int(page) * PAGE_SIZE

    if query.startswith('from:'):
        queryFrom(cur, query.replace('from:', "").strip(), offset)
    elif query.startswith('title:'):
        queryTitle(cur, query.replace('title:', "").strip(), offset)
    else:
        queryAll(cur, query, offset)

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

def queryFrom(cur, query, offset):
    if query is None:
        return
    log.info("query by sender")
    res = cur.execute("""
    Select Message_NormalizedSubject, Message_SenderList, Message_Preview, PathToDataFile, Message_TimeSent
        from Mail 
        where Message_SenderList LIKE ?
        ORDER BY Message_TimeSent DESC 
        LIMIT ? OFFSET ?
    """ , (unicode('%'+query+'%'), PAGE_SIZE, offset, ))

def queryTitle(cur, query, offset):
    if query is None:
        return
    log.info("query by subject")
    res = cur.execute("""
    Select Message_NormalizedSubject, Message_SenderList, Message_Preview, PathToDataFile, Message_TimeSent
        from Mail 
        where Message_NormalizedSubject LIKE ?
        ORDER BY Message_TimeSent DESC
        LIMIT ? OFFSET ?
    """ , (unicode('%'+query+'%'), PAGE_SIZE, offset, ))

def queryAll(cur, query, offset):
    if query is None:
        return
    log.info("query by subject, content and sender")
    res = cur.execute("""
    Select Message_NormalizedSubject, Message_SenderList, Message_Preview, PathToDataFile, Message_TimeSent
        from Mail 
        where 
        Message_NormalizedSubject LIKE ? 
        or Message_Preview LIKE ? 
        or Message_SenderList LIKE ?
        ORDER BY Message_TimeSent DESC
        LIMIT ? OFFSET ?
    """ , (unicode('%'+query+'%'), unicode('%'+query+'%'), unicode('%'+query+'%'), PAGE_SIZE, offset, ))

if __name__ == '__main__':
    wf = Workflow()
    log = wf.logger
    sys.exit(wf.run(main))