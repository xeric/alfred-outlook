#coding:utf-8
from __future__ import print_function, unicode_literals

import sys
import os
import re

from workflow import Workflow

log = None

KEY_PAGE_SIZE = 'pagesize'
ALL_KEYS = [KEY_PAGE_SIZE]

def main(wf):
    query = sys.argv[1]
    log.info(query)

    handle(query)

def handle(query):
    query = unicode(query, 'utf-8')

    if re.match(r'\S+\s+\S+', query) is None:
        return

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
            title='Set query page size to ' + val, 
            subtitle='', 
            arg=arg, 
            uid=arg, 
            valid=True
            )
    else:
        message = 'Unrecognize configuration name: ' + key

    wf.send_feedback()


if __name__ == '__main__':
    wf = Workflow()
    log = wf.logger
    sys.exit(wf.run(main))