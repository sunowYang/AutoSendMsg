#! coding=utf-8

import os
import sys
import traceback

from bin.ui import run
from bin.log import MyLog


# ********************************Get executing path******************************
if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(__file__)
# ********************************************************************************

log = MyLog(BASE_PATH, 'log.log')

if __name__ == '__main__':
    try:
        run(log, BASE_PATH)
    except Exception as e:
        log.logger.info(traceback.format_exc())
