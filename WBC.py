#! /usr/bin/env python2.7

# ----- Copyright (c) 2010-2017 by W. Craig Trader ---------------------------------
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published
# by the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

"""WBC: Generate iCal calendars from the WBC Schedule spreadsheet"""

# xxlint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0612,W0621,W0702,W0703
# pylint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0702

import logging
import sys

from WbcAllInOne import WbcAllInOne
from WbcMetadata import WbcMetadata
from WbcPreview import WbcPreview
from WbcScheduleComparison import ScheduleComparer
from WbcSpreadsheet import WbcSchedule

logging.basicConfig(level=logging.INFO)
logging.getLogger('requests').setLevel(logging.WARN)
LOG = logging.getLogger('WBC')

# ----- Real work happens here ------------------------------------------------

if __name__ == '__main__':
    meta = WbcMetadata()

    # Load a schedule from a spreadsheet, based upon commandline options.
    wbc_schedule = WbcSchedule(meta)
    wbc_schedule.create_all_calendars()

    # Parse the WBC Preview
    wbc_preview = WbcPreview(meta)

    # Parse the WBC All-in-One schedule
    wbc_allinone = WbcAllInOne(meta)

    # Compare the event calendars with the WBC All-in-One schedule and the preview
    comparer = ScheduleComparer(meta, wbc_schedule, wbc_allinone, wbc_preview)
    comparer.verify_event_calendars()

    LOG.info("Done.")
