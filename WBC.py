#! /usr/bin/env python2.7

#----- Copyright (c) 2010-2015 by W. Craig Trader ---------------------------------
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

from optparse import OptionParser
import logging

from WbcAllInOne import WbcAllInOne
from WbcMetadata import WbcMetadata
from WbcPreview import WbcPreview
from WbcScheduleComparison import ScheduleComparer
from WbcSpreadsheet import WbcSchedule

logging.basicConfig( level=logging.INFO )
LOGGER = logging.getLogger( 'WBC' )

#----- Time Constants --------------------------------------------------------

def process_options( metadata ):
    """
    Parse command line options
    """

    LOGGER.debug( 'Parsing commandline options' )

    parser = OptionParser()
    parser.add_option( "-y", "--year", dest="year", metavar="YEAR", default=metadata.this_year, help="Year to process" )
    parser.add_option( "-t", "--type", dest="type", metavar="TYPE", default="xls", help="Type of file to process (csv,xls)" )
    parser.add_option( "-i", "--input", dest="input", metavar="FILE", default=None, help="Schedule spreadsheet to process" )
    parser.add_option( "-o", "--output", dest="output", metavar="DIR", default="build", help="Directory for results" )
    parser.add_option( "-n", "--dry-run", dest="write_files", action="store_false", default=True )
    parser.add_option( "-v", "--verbose", dest="verbose", action="store_true", default=False )
    parser.add_option( "-f", "--full-report", dest="fullreport", action="store_true", default=False )
    options, dummy_args = parser.parse_args()

    options.year = int( options.year )

    if options.verbose:
        LOGGER.setLevel( logging.DEBUG )

    return options

#----- Real work happens here ------------------------------------------------

if __name__ == '__main__':

    meta = WbcMetadata()

    opts = process_options( meta )

    # Load a schedule from a spreadsheet, based upon commandline options.
    wbc_schedule = WbcSchedule( meta, opts )

    # Create calendar events from all of the spreadsheet events.
    wbc_schedule.create_wbc_calendars()

    if opts.write_files:
        # Write the individual event calendars.
        wbc_schedule.write_all_calendar_files()

        # Build the HTML index.
        wbc_schedule.write_index_page()

        # Output an improved copy of the input spreadsheet, in CSV
        wbc_schedule.write_spreadsheet()

    # Print the unmatched events for rework.
    wbc_schedule.report_unprocessed_events()

    # Parse the WBC All-in-One schedule
    wbc_allinone = WbcAllInOne( meta, opts )

    # Parse the WBC Preview
    names = dict( [( key_code, val_name ) for key_code, val_name in meta.names.items() ] )
    wbc_preview = WbcPreview( meta, opts, names )

    # Compare the event calendars with the WBC All-in-One schedule and the preview
    comparer = ScheduleComparer( meta, opts, wbc_schedule, wbc_allinone, wbc_preview )
    comparer.verify_event_calendars()

    LOGGER.warn( "Done." )
