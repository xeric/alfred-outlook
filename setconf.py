#coding:utf-8
from __future__ import print_function, unicode_literals

import sys
import re

from workflow import Workflow
from workflow.notify import notify
from consts import *

log = None


def main(wf):
    query = wf.decode(sys.argv[1])
    log.info(query)

    handle(query)


def handle(query):
    validQuery = False
    key = None
    val = None
    for rule in RULES:
        log.info(rule)
        m = re.match(rule, query)
        if m is not None:
            validQuery = True
            key = m.group(1)
            val = m.group(2)
            break
    if not validQuery:
        notify('Error', query + ' is not a valid configuration!')
        return

    if key is not None:
        wf.store_data(key, val)
        notify('Set Configuration Successfully', 'Set ' + key + " to " + val + " complete!")


if __name__ == '__main__':
    wf = Workflow()
    log = wf.logger
    sys.exit(wf.run(main))