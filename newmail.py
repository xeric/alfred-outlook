# encoding:utf-8
from __future__ import print_function, unicode_literals

import sys
import os
import re
import random

from workflow import Workflow
from workflow import Workflow3
from util import Util

GITHUB_SLUG = 'xeric/alfred-outlook'
UPDATE_FREQUENCY = 7

log = None

def main(wf):
    query = wf.decode(sys.argv[1])

    handle(wf, query)

    log.info('new mail composing')


def handle(wf, query):
    log.info(query)

    if len(query) < 1:
        wf.add_item(title='Type a Email Address Sent To',
                    subtitle='or chosse to Compose a Mail in Outlook Application',
                    arg='outlook ' + query,
                    uid=str(random.random()),
                    valid=True
                    )
    elif re.match('^message: ', query) and os.getenv('mail') is not None:
        content = re.sub(r'^message: ', '', query)
        receiver = os.getenv('mail')
        if len(content) > 0:
            wf.add_item(title='Send it to [' + receiver + ']',
                    subtitle='Send a quick message via mail',
                    arg='send: ' + receiver + " " + content,
                    uid=str(random.random()),
                    valid=True
                    )
        else:
            wf.add_item(title='Type a message after [message: ] to send ...',
                    subtitle='type a quick message to send',
                    arg='',
                    uid=str(random.random()),
                    valid=False
                    )
    elif not re.match('(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)', query): 
        wf.add_item(title='Email Address format is not valid',
                    subtitle='Type a valid mail address',
                    arg='',
                    uid=str(random.random()),
                    valid=False
                    )
    else:
        wf.add_item(title='Compose a Mail in Outlook Application',
                    subtitle='Outlook App window will be actived',
                    arg='outlook ' + query,
                    uid=str(random.random()),
                    valid=True
                    )
        if not Util.isAlfredV2(wf):
            it = wf.add_item(title='Send a Direct Quick Message ...',
                        subtitle='without activating Outlook App window',
                        arg='message: ',
                        uid=str(random.random()),
                        valid=True
                        )
            it.setvar('mail', query)

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
