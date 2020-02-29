# encoding:utf-8

import os
from consts import *

import workflow

APPLE_SCRIPT_DEFAULT_PROFILE = """
tell application "Microsoft Outlook"
	set theAlias to current identity folder as alias
	tell application "Finder"
		return name of theAlias
	end tell
end tell
"""


class Util(object):
    @staticmethod
    def isAlfredV2(wf):
        return wf.alfred_env['version'][0] == 2

    @staticmethod
    def configureDefaultProfile(wf):
        confFolder = workflow.util.run_applescript(APPLE_SCRIPT_DEFAULT_PROFILE)
        confFolder = confFolder.strip()
        # homePath = os.environ['HOME']
        #
        # defaultProfilePath = homePath + OUTLOOK_DATA_PARENT + OUTLOOK_DEFAULT_PROFILE
        # confFolder = OUTLOOK_DEFAULT_PROFILE
        #
        # if not os.path.isdir(defaultProfilePath):
        #     # list all folders
        #     parent = homePath + OUTLOOK_DATA_PARENT
        #     profileFolders = os.listdir(parent)
        #
        #     largestProfileSize = 0
        #     largestProfile = None
        #     for profile in profileFolders:
        #         if not profile.startswith('.'):
        #             if Util.validateProfile(parent + profile + OUTLOOK_DATA_FOLDER):
        #                 size = os.path.getsize(parent + profile + OUTLOOK_DATA_FOLDER + OUTLOOK_SQLITE_FILE)
        #                 if size > largestProfileSize:
        #                     largestProfile = profile
        #
        #     confFolder = largestProfile
        #
        # if confFolder:
        #     wf.logger.info("finally, we set profile to " + str(confFolder) + " as default profile conf!")
        # else:
        #     wf.logger.info("no profile found")
        wf.store_data(KEY_PROFILE, confFolder)

    @staticmethod
    def validateProfile(path):
        return os.path.isfile(path + OUTLOOK_SQLITE_FILE)