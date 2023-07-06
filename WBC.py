#! /usr/bin/env python3

# ----- Copyright (c) 2010-2022 by W. Craig Trader ---------------------------------
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

import logging

from WbcAllInOne import WbcAllInOne
from WbcCalendars import WbcWebcal
from WbcMetadata import WbcMetadata
from WbcPreview import WbcPreview
from WbcScheduleComparison import ScheduleComparer

logging.basicConfig(level=logging.INFO)
logging.getLogger('requests').setLevel(logging.WARN)
logging.getLogger('urllib3').setLevel(logging.WARN)

LOG = logging.getLogger('WBC')

# ----- Real work happens here ------------------------------------------------

if __name__ == '__main__':
    meta = WbcMetadata()

    # Create a calendar set
    wbc_calendars = WbcWebcal(meta)

    # Load a schedule from a spreadsheet, and populate the calendars, based upon commandline options.
    wbc_schedule = None
    if meta.type == 'old':
        from WbcOldSpreadsheet import WbcOldSchedule
        wbc_schedule = WbcOldSchedule(meta, wbc_calendars)
    elif meta.type == 'new':
        from WbcNewSpreadsheet import WbcNewSchedule
        wbc_schedule = WbcNewSchedule(meta, wbc_calendars)
    else:
        raise ValueError('Did not recognize spreadsheet type: %s' % meta.type)

    wbc_calendars.create_all_calendars()

    if meta.write_files:
        wbc_schedule.write_all_files()

    # Parse the WBC Preview
    wbc_preview = WbcPreview(meta)

    # Parse the WBC All-in-One schedule
    wbc_allinone = None # WbcAllInOne(meta)

    # Compare the event calendars with the WBC All-in-One schedule and the preview
    # comparer = ScheduleComparer(meta, wbc_calendars, wbc_allinone, wbc_preview)
    # comparer.verify_event_calendars()

    LOG.info("Done.")
