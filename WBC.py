#! /usr/bin/env python2.7

#----- Copyright (c) 2010-2016 by W. Craig Trader ---------------------------------
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

logging.basicConfig( level=logging.INFO )
logging.getLogger( 'requests' ).setLevel( logging.WARN )
LOG = logging.getLogger( 'WBC' )

#----- Real work happens here ------------------------------------------------

if __name__ == '__main__':

    meta = WbcMetadata()

    # Parse the WBC Preview
    # wbc_preview = WbcPreview( meta )
    # sys.exit(0)

    # Load a schedule from a spreadsheet, based upon commandline options.
    wbc_schedule = WbcSchedule( meta )

    # Create calendar events from all of the spreadsheet events.
    wbc_schedule.create_wbc_calendars()

    if meta.write_files:
        # Write the individual event calendars.
        wbc_schedule.write_all_calendar_files()

        # Build the HTML index.
        wbc_schedule.write_index_page()

        # Output an improved copy of the input spreadsheet, in CSV
        wbc_schedule.write_spreadsheet()

    # Print the unmatched events for rework.
    wbc_schedule.report_unprocessed_events()

    # Parse the WBC All-in-One schedule
    wbc_allinone = WbcAllInOne( meta )

    # Parse the WBC Preview
    wbc_preview = WbcPreview( meta )

    # Compare the event calendars with the WBC All-in-One schedule and the preview
    comparer = ScheduleComparer( meta, wbc_schedule, wbc_allinone, wbc_preview )
    comparer.verify_event_calendars()

    LOG.info( "Done." )
