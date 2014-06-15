#! /usr/bin/env python2.7

#----- Copyright (c) 2010-2014 by W. Craig Trader ---------------------------------
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

# These are the 'non-standard' libraries we need ...
#
# sudo apt-get install pip
# sudo pip install pytz
# sudo pip install beautifulsoup4
# sudo pip install icalendar
# sudo pip install xlrd
#
#

# xxlint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0612,W0621,W0702,W0703
# pylint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0702

from cgi import escape
from datetime import date, datetime, time, timedelta
from functools import total_ordering
from itertools import izip_longest
from optparse import OptionParser

import csv
import codecs
import logging
import os
import re
import shutil
import unicodedata
import urllib2

from bs4 import BeautifulSoup, Tag, NavigableString
from icalendar import Calendar, Event

import pytz
import xlrd

logging.basicConfig( level=logging.INFO )
LOGGER = logging.getLogger( 'WBC' )
DEBUGGING = True

#----- Time Constants --------------------------------------------------------

TZ = pytz.timezone( 'America/New_York' )  # Tournament timezone
UTC = pytz.timezone( 'UTC' )  # UTC timezone (for iCal)

#----- Utility methods -------------------------------------------------------

def parse_url( url ):
    """
    Utility function to load an HTML page from a URL, and parse it with BeautifulSoup.
    """

    page = None
    try:
        f = urllib2.urlopen( url )
        data = f.read()
        if ( len( data ) ):
            page = BeautifulSoup( data, "lxml" )
    except Exception as e:  # pylint: disable=W0703
        LOGGER.error( 'Failed while loading (%s)', url )
        LOGGER.error( e )

    return page

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

    return options

#----- WBC Event -------------------------------------------------------------

@total_ordering
class WbcRow( object ):
    """
    A WbcRow encapsulates information about a single schedule line from the 
    WBC schedule spreadsheet. This line may result in a dozen or more calendar events.
    """

    FIELDS = [ 'Date', 'Time', 'Event', 'PRIZE', 'CLASS', 'Format', 'Duration', 'Continuous', 'GM', 'Location', 'Code' ]

    def __init__( self, schedule, line, *args ):

        self.schedule = schedule
        self.meta = schedule.meta
        self.line = line

        # default values for calculated fields
        self.code = None
        self.name = None
        self.continuous = None
        self.duration = None
        self.length = None
        self.start = None
        self.datetime = None
        self.gm = None

        self.type = ''
        self.rounds = 0
        self.freeformat = False
        self.grognard = False
        self.junior = False

        # read the data row using the subclass
        self.readrow( *args )

        # This test is for debugging purposes; string search on the spreadsheet event name
        if DEBUGGING and self.name.find( 'Wits & Wagers' ) >= 0:
            pass

        # Check for errors that will throw exceptions later
        if self.gm == None:
            LOGGER.warning( 'Event "%s" missing gm', self.event )
            self.gm = ''

        if self.duration == None:
            LOGGER.warning( 'Event "%s" missing duration', self.event )
            self.duration = '0'

        if self.name.endswith( 'Final' ):
            self.name = self.name[:-4]

        # parse the data to generate useful fields
        self.checkrounds()
        self.checktypes( self.schedule.TYPES )
        self.checkrounds()
        self.checktypes( self.schedule.JUNIOR )
        self.checktimes()
        self.checkduration()
        self.checkcodes()

    @property
    def __key__( self ):
        return ( self.code, self.datetime, self.name )

    def __eq__( self, other ):
        return ( self.__key__ == other.__key__ )

    def __lt__( self, other ):
        return ( self.__key__ < other.__key__ )

    def __setattr__( self, key, value ):
        k = key.strip().lower()
        self.__dict__[ k ] = value

    def __getattr__( self, key ):
        k = key.strip().lower()
        if self.__dict__.has_key( k ):
            return self.__dict__[ k ]
        else:
            raise AttributeError( key )

    def __repr__( self ):
        return "%s @ %s %s on %s" % ( self.event, self.date, self.time, self.line )
#        return repr( self.__dict__ )

    @property
    def row( self ):
        row = dict( [ ( k, getattr( self, k ) ) for k in self.FIELDS ] )
        row['Date'] = self.date.strftime( '%Y-%m-%d' )
        row['Continuous'] = 'Y' if row['Continuous'] else ''
        return row

    def checkrounds( self ):
        """Check the current state of the event name to see if it describes a Heat or Round number"""

        match = re.search( r'([DHR]?)(\d+)[-/](\d+)$', self.name )
        if match:
            ( t, n, m ) = match.groups()
            text = match.group( 0 )
            text = text.replace( '-', '/' )
            if t == "R":
                self.start = int( n )
                self.rounds = int( m )
                self.name = self.name[:-len( text )].strip()
            elif t == "D":
                dtext = text.replace( 'D', '' )
                self.type = self.type + ' Demo ' + dtext
                self.type = self.type.strip()
                self.name = self.name[:-len( text )].strip()
            elif t == "H" or t == '':
                self.type = self.type + ' ' + text
                self.type = self.type.strip()
                self.name = self.name[:-len( text )].strip()

    def checktypes( self, types ):
        """
        Check the current state of the event name and strip off ( and flag ) any of the listed
        event type codes
        """

        for event_type in types:
            if self.name.endswith( event_type ):
                self.type = event_type + ' ' + self.type
                self.name = self.name[:-len( event_type )].strip()
                if event_type == 'FF':
                    self.freeformat = True
                elif event_type == 'PC':
                    self.grognard = True
                elif event_type in self.schedule.JUNIOR:
                    self.junior = True

        self.type = self.type.strip()

    def checkduration( self ):
        """
        Given the current event state, set the continuous event flag,
        and calculate the correct event length.
        """

        if self.__dict__.has_key( 'continuous' ):
            self.continuous = ( self.continuous in ( 'C', 'Y' ) )
        else:
            self.continuous = False

        if self.duration and self.duration.endswith( "q" ):
            self.continuous = True
            self.duration = self.duration[:-1]

        if self.duration and self.duration != '-':
            self.length = timedelta( minutes=60 * float( self.duration ) )

    def checkcodes( self ):
        """
        Check the current state of the event name and identify the actual event that matches
        the abbreviated name that's present.
        """

        self.code = None

        # First check for Junior events
        if self.junior:
            self.code = 'junior'

        # Check for tournament codes
        elif self.meta.codes.has_key( self.name ):
            self.code = self.meta.codes[ self.name ]
            self.name = self.meta.names[ self.code ]

            # If this event has rounds, save them for later use
            if self.rounds:
                self.schedule.rounds[ self.code ] = self.rounds

        else:
            # Check for non-tournament groupings
            for o in self.meta.others:
                if ( ( o['format'] and o['format'] == self.format ) or
                    ( o['name'] and o['name'] == self.name ) ):
                    self.code = o['code']
                    return

    def checktimes( self ):
        """Stub"""

        raise NotImplementedError()

    def readrow( self, *args ):
        """Stub"""

        raise NotImplementedError()

#----- WBC Event (read from CSV spreadsheet) ---------------------------------

class WbcCsvRow( WbcRow ):
    """This subclass of WbcRow is used to parse CSV-formatted schedule data"""

    def __init__( self, *args ):
        WbcRow.__init__( self, *args )

    def readrow( self, *args ):
        """Custom implementation of readrow to handle CSV-formatted rows"""

        row = args[0]

        for ( key, val ) in row.items():
            self.__setattr__( key, val )

        self.name = self.event.strip()

    def checktimes( self ):
        """Custom implementation of checktimes to handle CSV-formatted date/times"""

        try:
            d = datetime.strptime( self.date, "%m/%d/%Y" )
        except:
            try:
                d = datetime.strptime( self.date, "%m/%d/%y" )
            except:
                d = datetime.strptime( self.date, "%A, %B %d, %Y" )

        try:
            t = datetime.strptime( self.time, "%H" )
        except:
            try:
                t = datetime.strptime( self.time, "%H:%M" )
            except:
                t = datetime.strptime( self.time, "%I:%M:%S %p" )

        self.datetime = d.replace( hour=t.hour, minute=t.minute )

#----- WBC Event (read from Excel spreadsheet) -------------------------------

class WbcXlsRow( WbcRow ):
    """This subclass of WbcRow is used to parse Excel-formatted schedule data"""

    def __init__( self, *args ):
        WbcRow.__init__( self, *args )

    def readrow( self, *args ):
        """Custom implementation of readrow to handle XLS-formatted rows"""

        labels = args[0]
        row = args[1]
        datemode = args[2]

        for i in range( len( labels ) ):
            key = labels[i]

            val = row[i]
            if not key:
                continue
            elif val.ctype == xlrd.XL_CELL_EMPTY:
                val = None
            elif val.ctype == xlrd.XL_CELL_TEXT:
                val = unicodedata.normalize( 'NFKD', val.value ).encode( 'ascii', 'ignore' ).strip()
            elif val.ctype == xlrd.XL_CELL_NUMBER:
                val = str( float( val.value ) )
            elif val.ctype == xlrd.XL_CELL_DATE:
                val = xlrd.xldate_as_tuple( val.value, datemode )
                if val[0]:
                    val = datetime( *val )  # pylint: disable=W0142
                else:
                    val = time( val[3], val[4], val[5] )
            else:
                raise ValueError( "Unhandled Excel cell type (%s) for %s" % ( val.ctype, key ) )

            self.__setattr__( key, val )

        self.name = self.event.strip()

    def checktimes( self ):
        """Custom implementation of checktimes to handle XLS-formatted date/times"""

        d = self.date

        if self.time.__class__ is time:
            t = self.time
        else:
            try:
                t = float( self.time )
                if t > 23:
                    t = time( 23, 59 )
                    self.length = timedelta( minutes=1 )
                else:
                    m = t * 60
                    h = int( m / 60 )
                    m = int( m % 60 )
                    t = time( h, m )
            except:
                try:
                    t = datetime.strptime( self.time, "%H:%M" )
                except:
                    try:
                        t = datetime.strptime( self.time, "%I:%M:%S %p" )
                    except:
                        raise ValueError( 'Unable to format (%s) as a time' % self.time )

        self.datetime = d.replace( hour=t.hour, minute=t.minute )

#----- WBC Meta Data ---------------------------------------------------------

class WbcMetadata( object ):
    """Load metadata about events that is not available from other sources"""

    this_year = datetime.now( TZ ).year

    # Data file names
    EVENTCODES = "wbc-event-codes.csv"
    OTHERCODES = "wbc-other-codes.csv"

    others = []  # List of non-tournament event matching data
    special = []  # List of non-tournament event codes
    tourneys = []  # List of tournament codes

    codes = {}  # Name -> Code map for events
    names = {}  # Code -> Name map for events

    durations = {}  # Special durations for events that have them
    grognards = {}  # Special durations for grognard events that have them
    playlate = {}  # Flag for events that may run past midnight

    first_day = None  # First calendar day for this year's convention

    def __init__( self ):
        self.load_tourney_codes()
        self.load_other_codes()

    def load_tourney_codes( self ):
        """
        Load all of the tourney codes (and alternate names) from their data file.
        """

        LOGGER.debug( 'Loading tourney event codes' )

        codefile = csv.DictReader( open( self.EVENTCODES ) )
        for row in codefile:
            c = row['Code'].strip()
            n = row['Name'].strip()
            self.codes[ n ] = c
            self.names[ c ] = n
            self.tourneys.append( c )

            if row['Duration']:
                self.durations[c] = int( row['Duration'] )

            if row['Grognard']:
                self.grognards[c] = int( row['Grognard'] )

            if row['PlayLate']:
                self.playlate[c] = row['PlayLate'].strip().lower()

            for altname in [ 'Alt1', 'Alt2', 'Alt3', 'Alt4', 'Alt5', 'Alt6']:
                if row[altname]:
                    a = row[altname].strip()
                    self.codes[a] = c

    def load_other_codes( self ):
        """
        Load all of the non-tourney codes from their data file.
        """

        LOGGER.debug( 'Loading non-tourney event codes' )

        codefile = csv.DictReader( open( self.OTHERCODES ) )
        for row in codefile:
            c = row['Code'].strip()
            d = row['Description'].strip()
            n = row['Name'].strip()
            f = row['Format'].strip()

            other = { 'code' : c, 'description' : d, 'name' : n, 'format' : f }
            self.others.append( other )
            self.special.append( c )
            self.names[ c ] = d

    def check_date( self, event_date ):
        """Check to see if this event date is the earliest event date seen so far"""
        if not self.first_day:
            self.first_day = event_date
        elif event_date < self.first_day:
            self.first_day = event_date

#----- WBC Schedule ----------------------------------------------------------

class WbcSchedule( object ):
    """
    The WbcSchedule class parses the entire WBC schedule spreadsheet and creates 
    iCalendar calendars for each event (with vEvents for each time slot).
    """
    valid = False

    # Data file names
    TEMPLATE = "index-template.html"

    # Recognized event flags
    FLAVOR = [ 'AFC', 'NFC', 'FF', 'Circus', 'DDerby', 'Draft', 'Playoffs', 'FF' ]
    JUNIOR = [ 'Jr', 'Jr.', 'Junior' ]
    TEEN = [ 'Teen' ]
    MENTOR = [ 'Mentoring' ]
    MULTIPLE = ['QF/SF/F', 'QF/SF', 'SF/F' ]
    SINGLE = [ 'QF', 'SF', 'F' ]
    STYLE = [ 'After Action Debriefing', 'After Action', 'Aftermath', 'Awards', 'Demo', 'Mulligan' ] + MULTIPLE + SINGLE

    TYPES = [ 'PC' ] + FLAVOR + JUNIOR + TEEN + MENTOR + STYLE

    rounds = {}  # Number of rounds for events that have rounds
    events = {}  # Events, grouped by code and then sorted by start date/time
    unmatched = []  # List of spreadsheet rows that don't match any coded events
    calendars = {}  # Calendars for each event code
    locations = {}  # Calendars by location
    dailies = {}  # Calendars by date

    current_tourneys = []
    everything = None
    tournaments = None
    options = None

    def __init__( self, metadata, options ):
        """
        Initialize a schedule
        """
        self.processed = datetime.now( TZ )

        self.meta = metadata
        self.options = options

        self.year = self.options.year

        if not os.path.exists( self.options.output ):
            os.makedirs( self.options.output )

        self.load_events()

        self.prodid = "WBC %s" % self.options.year

        self.valid = True

    def load_events( self ):
        """
        Process all of the events in the spreadsheet
        """

        LOGGER.info( 'Scanning schedule spreadsheet' )

        filename = self.options.input
        if self.options.type == 'csv':
            if not filename:
                filename = "wbc-%s-schedule.csv" % self.options.year
            self.scan_csv_file( filename )
        elif self.options.type == 'xls':
            if not filename:
                filename = "schedule%s.xls" % self.options.year
            self.scan_xls_file( filename )

    def scan_csv_file( self, filename ):
        """
        Read a CSV-formatted file and generate WBC events for each row
        """

        LOGGER.debug( 'Reading CSV spreadsheet from %s', filename )

        with open( filename ) as f:
            eventfile = csv.DictReader( f , delimiter=';' )
            i = 2
            for row in eventfile:
                event_row = WbcCsvRow( self, i, row )
                self.categorize_row( event_row )
                i = i + 1

    def scan_xls_file( self, filename ):
        """
        Read an Excel spreadsheet and generate WBC events for each row.
        """

        LOGGER.debug( 'Reading Excel spreadsheet from %s', filename )

        book = xlrd.open_workbook( filename )
        sheet = book.sheet_by_index( 0 )

        header = []
        for i in range( sheet.ncols ):
            try:
                key = sheet.cell_value( 0, i )
                if key:
                    key = unicodedata.normalize( 'NFKD', key ).encode( 'ascii', 'ignore' ).lower()
                header.append( key )
            except:
                raise ValueError( 'Unable to parse Column Header %d (%s)' % ( i, key ) )

        for i in range( 1, sheet.nrows ):
            if self.options.verbose:
                LOGGER.debug( 'Reading row %d' % ( i + 1 ) )
            event_row = WbcXlsRow( self, i + 1, header, sheet.row( i ), book.datemode )
            self.categorize_row( event_row )

    def categorize_row( self, row ):
        """
        Assign a spreadsheet entry to a matching row code.
        """
        if row.code:
            if self.events.has_key( row.code ):
                self.events[row.code].append( row )
            else:
                self.events[row.code] = [ row ]
        else:
            self.unmatched.append( row )

        self.meta.check_date( row.date )

    def create_wbc_calendars( self ):
        """
        Process all of the spreadsheet entries, by event code, then by time,
        creating calendars for each entry as needed.
        """

        LOGGER.info( 'Creating calendars' )

        for event_list in self.events.values():
            event_list.sort( lambda x, y: cmp( x.datetime, y.datetime ) )
            for event in event_list:
                self.process_event( event )

        # Create a sorted list of this year's tourney codes
        self.current_tourneys = [ code for code in self.calendars.keys()
                                       if code in self.meta.tourneys ]
        self.current_tourneys.sort( lambda x, y: cmp( self.meta.names[x], self.meta.names[y] ) )

        # Create bulk calendars
        self.everything = Calendar()
        self.everything.add( 'VERSION', '2.0' )
        self.everything.add( 'PRODID', '-//' + self.prodid + ' Everything//ct7//' )
        self.everything.add( 'SUMMARY', 'WBC %s All-in-One Schedule' % self.options.year )

        self.tournaments = Calendar()
        self.tournaments.add( 'VERSION', '2.0' )
        self.tournaments.add( 'PRODID', '-//' + self.prodid + ' Tournaments//ct7//' )
        self.tournaments.add( 'SUMMARY', 'WBC %s Tournaments Schedule' % self.options.year )

        # For all of the event calendars
        for code, calendar in self.calendars.items():

            # Add all calendar events to the master calendar
            self.everything.subcomponents += calendar.subcomponents

            # Add all the tourney events to the tourney calendar
            if code in self.current_tourneys:
                self.tournaments.subcomponents += calendar.subcomponents

            # For each calendar event
            for event in calendar.subcomponents:

                # Add it to the appropriate location calendar
                location = self.get_or_create_location_calendar( event['LOCATION'] )
                location.subcomponents.append( event )

                # Add it to the appropriate daily calendar
                daily = self.get_or_create_daily_calendar( event['DTSTART'] )
                daily.subcomponents.append( event )

    def process_event( self, event ):
        """
        For a spreadsheet entry, generate calendar events as follows:
        
        If the event is WAW, 
            treat it like an all-week free-format event.
        If the event is free-format, and a grognard,
            it's an all-week event, but use the grognard duration from the event codes.
        If the event is free-format, and a Swiss Elimination, and it has rounds,
            it's an all-week event, but use the duration from the event codes.
        If the event has rounds, 
            add calendar events for each round.
        If the event is marked as continuous, and it's marked 'HMSE', 
           there's no clue as to how many actual heats there are, so just add one event.
        If the event is marked as continuous, 
            add as many events as there are types coded.
        Otherwise, 
           add a single event for the event.
        """

        calendar = self.get_or_create_event_calendar( event.code )

        # This test is for debugging purposes, and is only good for an event that was sucessfully coded
        if DEBUGGING and event.code in ( 'LST', ):
            pass

        if event.code == 'WAW':
            self.process_all_week_event( calendar, event )
        elif event.freeformat and event.grognard:
            self.process_freeformat_grognard_event( calendar, event )
        elif event.freeformat and event.format == 'SwEl' and event.rounds:
            self.process_freeformat_swel_event( calendar, event )
        elif event.rounds:
            self.process_event_with_rounds( calendar, event )
        elif event.continuous and event.format == 'HMSE':
            self.process_normal_event( calendar, event )
        elif event.continuous:
            self.process_continuous_event( calendar, event )
        else:
            self.process_normal_event( calendar, event )

    def process_normal_event( self, calendar, event ):
        """
        Process a spreadsheet entry that maps to a single event.
        """
        name = event.name + ' ' + event.type
        alternative = self.alternate_round_name( event )
        if alternative:
            self.add_event( calendar, event, name=name, altname=alternative )
        else:
            self.add_event( calendar, event, name=name, replace=False )

    def process_continuous_event( self, calendar, entry ):
        """
        Process multiple back-to-back events that are not rounds, per se.
        """
        start = entry.datetime
        for event_type in entry.type.split( '/' ):
            name = entry.name + ' ' + event_type
            alternative = self.alternate_round_name( entry, event_type )
            self.add_event( calendar, entry, start=start, name=name, altname=alternative )
            start = self.calculate_next_start_time( entry, start )

    def process_event_with_rounds( self, calendar, entry ):
        """
        Process multiple back-to-back rounds
        """
        start = entry.datetime
        name = entry.name + ' ' + entry.type
        name = name.strip()

        rounds = range( int( entry.start ), int( entry.rounds ) + 1 )
        for r in rounds:
            label = "%s R%s/%s" % ( name, r, entry.rounds )
            self.add_event( calendar, entry, start=start, name=label )
            start = self.calculate_next_start_time( entry, start )

    def calculate_next_start_time( self, entry, start ):
        """
        Calculate when to start the next round.
        
        In theory, WBC runs from 9am to midnight.  Thus rounds that would 
        otherwise begin or end after midnight should be postponed until 
        9am the next morning.  There are three types of exceptions to this
        rule:  
        
        (1) some events run their finals directly after the semi-finals conclude,
            which can cause the final to run past midnight.
        (2) some events are scheduled to run multiple rounds past midnight.
        (3) some events are variable -- starting a 6 hour round at 9pm is OK, but not at 10pm
        """

        # Calculate midnight, relative to the last event's start time
        midnight = datetime.fromordinal( start.toordinal() + 1 )

        # Calculate 9am tomorrow
        tomorrow = datetime( midnight.year, midnight.month, midnight.day, 9, 0, 0 )

        # Nominal start time and end time for the next event
        next_start = start + entry.length
        next_end = next_start + entry.length

        late_part = 0 if next_end <= midnight else ( next_end - midnight ).total_seconds() / entry.length.total_seconds()

        # Lookup the override code for this event
        playlate = self.meta.playlate.get( entry.code, None )

        if playlate and late_part:
            LOGGER.warn( "%s: %4s, Start: %s, End: %s, Partial: %5.2f, %s", entry.code, playlate, next_start, next_end, late_part, late_part <= 0.5 )

        if playlate == 'all':
            pass  # always
        elif next_start > midnight:
            next_start = tomorrow
        elif next_end <= midnight:
            pass
        elif playlate == 'once' and late_part <= 0.5:
            pass
        else:
            next_start = tomorrow

        return next_start

    def process_freeformat_swel_event( self, calendar, entry ):
        """
        Process an entry that is a event with no fixed schedule.
        
        These events run continuously for several days, followed by 
        separate semi-final and finals.  This is the same as an 
        all-week-event, except that the event duration and name are wrong.
        """

        duration = self.meta.durations[entry.code] if self.meta.durations.has_key( entry.code ) else 51
        label = "%s R%s/%s" % ( entry.name, 1, entry.rounds )
        self.process_all_week_event( calendar, entry, duration, label )

    def process_freeformat_grognard_event( self, calendar, entry ):
        """
        Process an entry that is a pre-con event with no fixed schedule.
        
        These events run for 10 hours on Saturday, 15 hours on Sunday, 
        15 hours on Monday, and 9 hours on Tuesday, before switching to a 
        normal tourney schedule.  This is the same as an all-week-event, 
        except that the event duration and name are wrong.
          
        In this case, the duration is 10 + 15 + 15 + 9 = 49 hours. 
        """

        duration = self.meta.grognards[entry.code] if self.meta.grognards.has_key( entry.code ) else 49
        label = "%s PC R%s/%s" % ( entry.name, 1, entry.rounds )
        self.process_all_week_event( calendar, entry, duration, label )

    def process_all_week_event( self, calendar, entry, length=None, label=None ):
        """
        Process an entry that runs continuously all week long.
        """

        start = entry.datetime
        remaining = timedelta( hours=length ) if length else entry.length
        label = label if label else entry.name

        while ( remaining.days or remaining.seconds ):
            midnight = start.date() + timedelta( days=1 )
            duration = datetime( midnight.year, midnight.month, midnight.day ) - start
            if duration > remaining:
                duration = remaining

            self.add_event( calendar, entry, start=start, duration=duration, replace=False, name=label )

            start = datetime( midnight.year, midnight.month, midnight.day, 9, 0, 0 )
            remaining = remaining - duration

    def alternate_round_name( self, entry, event_type=None ):
        """
        Create the equivalent round name for a given entry.
        
        At WBC, a tournament is typically composed a set of elimination rounds.
        The first round can be composed of multiple qualification heats, and 
        a tournament may also include a mulligan round.  After that, rounds are
        identified by the round number (eg: R2/4 represents Round 2 of 4).
        To make matters more complicated, the last round of a tournament is 
        typically labeled 'F' for Final, and the next-to-last round would be
        'SF' for SemiFinal.  For a large enough tournament, there may even be
        a quarter-final round.  Thus it is possible to see a tourney with multiple 
        heats (H1, H2, H3), and then R2/6, R3/6, QF, SF, F, where R4/6 = QF, 
        R5/6 = SF, and R6/6 = F.
        
        This method will calculate what the generic round name should be, 
        if the current entry type is 'QF', 'SF', or 'F'.   
        """

        event_type = event_type if event_type else entry.type

        alternative = None
        if self.rounds.has_key( entry.code ) and event_type in self.SINGLE:
            r = self.rounds[ entry.code ]
            offset = ( len( self.SINGLE ) - self.SINGLE.index( event_type ) ) - 1
            alternative = "%s R%s/%s" % ( entry.name, r - offset, r )
        return alternative

    def get_or_create_event_calendar( self, code ):
        """
        For a given event code, return the iCalendar that matches that code.
        If there is no pre-existing calendar, create a new one.
        """
        if self.calendars.has_key( code ):
            return self.calendars[ code ]

        description = "%s %s: %s" % ( self.prodid, code, self.meta.names[code] )
        url = 'http://boardgamers.org/yearbkex/%spge.htm' % code.lower()

        calendar = Calendar()
        calendar.add( 'VERSION', '2.0' )
        calendar.add( 'PRODID', '-//%s %s//ct7//' % ( self.prodid, code ) )
        calendar.add( 'SUMMARY', self.meta.names[ code ] )
        calendar.add( 'DESCRIPTION', description )
        calendar.add( 'URL', url )

        self.calendars[ code ] = calendar

        return calendar

    def get_or_create_location_calendar( self, location ):
        """
        For a given location, return the iCalendar that matches that location.
        If there is no pre-existing calendar, create a new one.
        """
        location = str( location ).strip()
        if self.locations.has_key( location ):
            return self.locations[ location ]

        description = "%s: Events in %s" % ( self.prodid, location )

        calendar = Calendar()
        calendar.add( 'VERSION', '2.0' )
        calendar.add( 'PRODID', '-//%s %s//ct7//' % ( self.prodid, location ) )
        calendar.add( 'SUMMARY', 'Events in ' + location )
        calendar.add( 'DESCRIPTION', description )

        self.locations[ location ] = calendar

        return calendar

    def get_or_create_daily_calendar( self, event_date ):
        """
        For a given date, return the iCalendar that matches that date.
        If there is no pre-existing calendar, create a new one.
        """
        key = event_date.dt.date()
        name = event_date.dt.strftime( '%A, %B %d' )

        if self.dailies.has_key( key ):
            return self.dailies[ key ]

        description = '%s: Events on %s' % ( self.prodid, name )

        calendar = Calendar()
        calendar.add( 'VERSION', '2.0' )
        calendar.add( 'PRODID', '-//%s %s//ct7//' % ( self.prodid, key ) )
        calendar.add( 'SUMMARY', 'Events on ' + name )
        calendar.add( 'DESCRIPTION', description )

        self.dailies[ key ] = calendar

        return calendar

    def add_event( self, calendar, entry, start=None, duration=None, name=None, altname=None, replace=True ):
        """
        Add a new vEvent to the given iCalendar for a given spreadsheet entry.
        """
        name = name if name else entry.name
        start = start if start else entry.datetime
        duration = duration if duration else entry.length

        localized_start = TZ.localize( start )
        utc_start = localized_start.astimezone( UTC )

        e = Event()
        e.add( 'SUMMARY', name )
        e.add( 'DTSTART', utc_start )
        e.add( 'DURATION', duration )
        e.add( 'LOCATION', entry.location )
        e.add( 'CONTACT', entry.gm )

        if replace:
            self.add_or_replace_event( calendar, e, altname )
        else:
            calendar.add_component( e )

    def add_or_replace_event( self, calendar, event, altname=None ):
        """
        Insert a vEvent into an iCalendar.  
        If the vEvent 'matches' an existing vEvent, replace the existing vEvent instead.  
        """
        for i in range( len( calendar.subcomponents ) ):
            if self.is_same_icalendar_event( calendar.subcomponents[i], event, altname ):
                calendar.subcomponents[i] = event
                return
        calendar.subcomponents.append( event )

    def report_unprocessed_events( self ):
        """
        Report on all of the WBC schedule entries that were not processed.
        """

        self.unmatched.sort( cmp=lambda x, y: cmp( x.name, y.name ) )
        for event in self.unmatched:
            LOGGER.error( 'Did not process Row %3d [%s] %s', event.line, event.name, event )

    def write_calendar_file( self, calendar, name ):
        """
        Write an actual calendar file, using a filesystem-safe name.
        """
        filename = self.safe_ics_filename( name )
        with open( os.path.join( self.options.output, filename ), "wb" ) as f:
            f.write( self.serialize_calendar( calendar ) )

    def write_all_calendar_files( self ):
        """
        Write all of the calendar files.
        """
        LOGGER.info( "Saving calendars..." )

        # Remote any existing destination directory
        if os.path.exists( self.options.output ):
            shutil.rmtree( self.options.output )

        # Create the destination directory
        os.makedirs( self.options.output )

        # Copy needed files to the destination
        if os.path.exists( 'ical.gif' ):
            shutil.copy( 'ical.gif', self.options.output )

        # For all of the event calendars
        for code, calendar in self.calendars.items():

            # Write the calendar itself
            self.write_calendar_file( calendar, code )

        # Write the master and tourney calendars
        self.write_calendar_file( self.everything, "all-in-one" )
        self.write_calendar_file( self.tournaments, "tournaments" )

        # Write the location calendars
        for location, calendar in self.locations.items():
            self.write_calendar_file( calendar, location )

        # Write the daily calendars
        for day, calendar in self.dailies.items():
            self.write_calendar_file( calendar, day )

    def write_spreadsheet( self ):
        """
        Write all of the calendar entries back out, in CSV format, with improvements
        """
        LOGGER.info( 'Writing spreadsheet...' )

        data = []
        for k in self.events.keys():
            data = data + self.events[k]
        data = data + self.unmatched
        data.sort()


        with open( os.path.join( self.options.output, "schedule.csv" ), "w" ) as f:
            writer = csv.DictWriter( f, WbcRow.FIELDS, extrasaction='ignore' )
            writer.writeheader()
            writer.writerows( [ e.row for e in data ] )

    def write_index_page( self ):
        """
        Using an HTML Template, create an index page that lists 
        all of the created calendars.
        """

        LOGGER.info( 'Writing index page...' )

        with open( self.TEMPLATE, "r" ) as f:
            template = f.read()

        parser = BeautifulSoup( template )

        # Locate insertion points
        title = parser.find( 'title' )
        header = parser.find( 'div', { 'id' : 'header' } )
        footer = parser.find( 'div', { 'id' : 'footer' } )

        # Page title
        title.insert( 0, parser.new_string( "WBC %s Event Schedule" % self.options.year ) )
        header.h1.insert( 0, parser.new_string( "WBC %s Event Schedule" % self.options.year ) )
        footer.p.insert( 0, parser.new_string( "Updated on %s" % self.processed.strftime( "%A, %d %B %Y %H:%M %Z" ) ) )

        # Tournament event calendars
        tourneys = dict( [( k, v ) for k, v in self.calendars.items() if k not in self.meta.special ] )
        ordering = lambda x, y: cmp( tourneys[x]['summary'], tourneys[y]['summary'] )
        self.render_calendar_table( parser, 'tournaments', 'Tournament Events', tourneys, ordering )

        # Non-tourney event calendars
        nontourneys = dict( [( k, v ) for k, v in self.calendars.items() if k in self.meta.special ] )
        self.render_calendar_list( parser, 'other', 'Other Events', nontourneys )

        # Location calendars
        self.render_calendar_list( parser, 'location', 'Location Calendars', self.locations )

        # Daily calendars
        self.render_calendar_list( parser, 'daily', 'Daily Calendars', self.dailies )

        # Special event calendars
        specials = {}
        specials['all-in-one'] = self.everything
        specials['tournaments'] = self.tournaments
        self.render_calendar_list( parser, 'special', 'Special Calendars', specials )

        with codecs.open( os.path.join( self.options.output, 'index.html' ), 'w', 'utf-8' ) as f:
            f.write( parser.prettify() )

    @classmethod
    def render_calendar_table( cls, parser, id_name, label, calendar_map, comparison=None ):
        """Create the HTML fragment for the table of tournament calendars."""

        keys = calendar_map.keys()
        keys.sort( comparison )

        div = parser.find( 'div', { 'id' : id_name } )
        div.insert( 0, parser.new_tag( 'h2' ) )
        div.h2.insert( 0, parser.new_string( label ) )
        div.insert( 1, parser.new_tag( 'table' ) )

        for row_keys in cls.split_list( keys, 2 ):
            tr = parser.new_tag( 'tr' )
            div.table.insert( len( div.table ), tr )

            for key in row_keys:
                label = calendar_map[ key ]['summary'] if key else ''
                td = cls.render_calendar_table_entry( parser, key, label )
                tr.insert( len( tr ), td )

    @classmethod
    def render_calendar_table_entry( cls, parser, key, label ):
        """Create the HTML fragment for one cell in the tournament calendar table."""
        td = parser.new_tag( 'td' )
        if key:
            span = parser.new_tag( 'span' )
            span['class'] = 'eventcode'
            span.insert( 0, parser.new_string( escape( key ) + ': ' ) )
            td.insert( len( td ), span )

            filename = cls.safe_ics_filename( key )

            a = parser.new_tag( 'a' )
            a['class'] = 'eventlink'
            a['href'] = '#'
            a['onclick'] = "webcal('%s');" % filename
            img = parser.new_tag( 'img' )
            img['src'] = 'ical.gif'
            a.insert( len( a ), img )
            td.insert( len( td ), a )

            a = parser.new_tag( 'a' )
            a['class'] = 'eventlink'
            a['href'] = filename
            a.insert( 0, parser.new_string( escape( "%s" % label ) ) )
            td.insert( len( td ), a )

            td.insert( len( td ), a )
        else:
            td.insert( len( td ) , parser.new_string( '&nbsp;' ) )
        return td

    @staticmethod
    def split_list( original, width ):
        """
        A generator that, given a list of indeterminate length, will split the list into 
        roughly equal columns, and then return the resulting list one row at a time. 
        """

        max_length = len( original )
        length = ( max_length + width - 1 ) / width
        for i in range( length ):
            partial = []
            for j in range( width ):
                k = i + j * length
                partial.append( original[k] if k < max_length else None )
            yield partial

    @classmethod
    def render_calendar_list( cls, parser, id_name, label, calendar_map, comparison=None ):
        """Create the HTML fragment for an unordered list of calendars."""

        keys = calendar_map.keys()
        keys.sort( comparison )

        div = parser.find( 'div', { 'id' : id_name } )
        div.insert( 0, parser.new_tag( 'h2' ) )
        div.h2.insert( 0, parser.new_string( label ) )
        div.insert( 1, parser.new_tag( 'ul' ) )

        for key in keys:
            calendar = calendar_map[ key ]
            cls.render_calendar_list_item( parser, div.ul, key, calendar['summary'] )

    @classmethod
    def render_calendar_list_item( cls, parser, list_tag, key, label ):
        """Create the HTML fragment for a single calendar in a list"""

        li = parser.new_tag( 'li' )

        filename = cls.safe_ics_filename( key )

        a = parser.new_tag( 'a' )
        a['class'] = 'eventlink'
        a['href'] = '#'
        a['onclick'] = "webcal('%s');" % filename
        img = parser.new_tag( 'img' )
        img['src'] = 'ical.gif'
        a.insert( len( a ), img )
        li.insert( len( li ), a )

        a = parser.new_tag( 'a' )
        a['class'] = 'eventlink'
        a['href'] = filename
        a.insert( 0, parser.new_string( escape( "%s" % label ) ) )
        li.insert( len( li ), a )

        list_tag.insert( len( list_tag ), li )

    @classmethod
    def serialize_calendar( cls, calendar ):
        """This fixes portability quirks in the iCalendar library:
        1) The iCalendar library generates event start date/times as 'DTSTART;DATE=VALUE:yyyymmddThhmmssZ';
           the more acceptable format is 'DTSTART:yyyymmddThhmmssZ'
        2) The iCalendar library doesn't sort the events in a given calendar by date/time.
        """

        output = calendar.to_ical()
        output = output.replace( ";TZID=UTC;VALUE=DATE-TIME:", ":" )
        return output

    @staticmethod
    def compare_icalendar_events( x, y ):
        """
        Comparison method for iCal events
        """
        c = cmp( x['dtstart'].dt, y['dtstart'].dt )
        c = cmp( x['summary'], y['summary'] ) if not c else c
        return c

    @staticmethod
    def is_same_icalendar_event( e1, e2, altname=None ):
        """
        Compare two events to determine if they are 'the same'.
        
        If they start at the same time, and have the same duration, they're 'the same'.
        If they have the same name, they're 'the same'. 
        If the first matches an alternative name, they're 'the same'.
        """
        same = str( e1['dtstart'] ) == str( e2['dtstart'] )
        same &= str( e1['duration'] ) == str( e2['duration'] )
        same |= unicode( e1['summary'] ) == unicode( e2['summary'] )
        if altname:
            same |= unicode( e1['summary'] ) == unicode( altname )
        return same

    @staticmethod
    def safe_ics_filename( name ):
        """
        Given an object, determine a web-safe filename from it, then append '.ics'.
        """
        if name.__class__ is date:
            name = name.strftime( "%Y-%m-%d" )
        else:
            name = name.strip()
            name = name.replace( '&', 'n' )
            name = name.replace( ' ', '_' )
            name = name.replace( '/', '_' )
        return "%s.ics" % name

#----- WBC All-in-One Schedule -----------------------------------------------

class WbcAllInOne( object ):
    """
    This class is used to parse the published All-in-One Schedule, and produce 
    a list of tourney events that can be used to compare against the calendars 
    generated by the WbcSchedule class.  
    
    The comparer is really just a sanity check, because there is less 
    information present in the All-in-One Schedule than is needed to build a 
    correct calendar entry.  On the other hand, it's easier to parse than the
    YearBook pages for each event.

    A typical row on the All-in-One schedule might look something like this:
        
        <tr><td>
        <i><FONT SIZE=+2>RBS</FONT></i>
        </td><td align=right valign=top bgcolor="#FFFF00">
        <FONT COLOR="#000000"><i>Russian Beseiged</i>
        </FONT>
        </td><td>
        </td><td>
        &nbsp</td><td>
        &nbsp</td><td>
        &nbsp</td><td>
        &nbsp</td><td>
        We<FONT COLOR=green>17</FONT>,<FONT COLOR=magenta>19</FONT><br>
        <FONT SIZE=-1>17:Pt; 19:Lampeter</FONT>
        </td><td>
        Th<FONT COLOR=red>9</FONT>,<FONT COLOR=red>14</FONT>,<FONT COLOR=blue>19</FONT><br>
        <FONT SIZE=-1>Lampeter</FONT>
        </td><td>
        Fr<FONT COLOR=#AAAA00>9</FONT><br>
        <FONT SIZE=-1>Lampeter</FONT>
        </td><td>
        &nbsp</td><td>
        &nbsp</tr>
        
    This is, frankly, horrible HTML.  But at least it's consistent, year-to-year, 
    and BeautifulSoup can parse it. 
    """

    SITE_URL = 'http://boardgamers.org/wbc/allin1.htm'

    valid = False

    TERRACE = 'Pt'

    colormap = {
        'green': 'Demo',
        '#07BED2': 'Mulligan',
        'red': 'Round',
        'blue': 'SF',
        '#AAAA00': 'F',
    }

    # Events that are miscoded (bad code : actual code)
    codemap = { 'MMA': 'MRA', }

    events = {}

    class Event( object ):
        """Simple data object to collect information about an event occuring at a specific time."""

        def __init__( self ):
            self.code = None
            self.name = None
            self.type = None
            self.time = None
            self.location = None

        def __cmp__( self, other ):
            return cmp( self.time, other.time )

        def __str__( self ):
            return '%s %s %s in %s at %s' % ( self.code, self.name, self.type, self.location, self.time )

    def __init__( self, metadata, options ):
        self.meta = metadata
        self.options = options
        self.page = None

        self.load_table()

    def load_table( self ):
        """Parse the All-in-One schedule (HTML)"""

        LOGGER.info( 'Parsing WBC All-in-One schedule' )

        self.page = parse_url( self.SITE_URL )
        if not self.page:
            return

        try:
            title = self.page.findAll( 'title' )[0]
            year = str( title.text )
            year = year.strip().split()
            year = int( year[0] )
        except:
            # Fetch from page body instead of page title.
            # html.body.table.tr.td.p.font.b.font.NavigableString
            try:
                text = self.page.html.body.table.tr.td.p.font.b.font.text
                year = str( text ).strip().split()
                year = int( year[0] )
            except:
                year = 2013

        if year != self.meta.this_year and year != self.options.year:
            LOGGER.error( "All-in-one schedule for %d is out of date", year )

            return

        tables = self.page.findAll( 'table' )
        rows = tables[1].findAll( 'tr' )
        for row in rows[1:]:
            self.load_row( row )

        self.valid = True

    def load_row( self, row ):
        """Parse an individual all-in-one row to find times and rooms for an event"""

        events = []

        cells = row.findAll( 'td' )
        code = str( cells[0].font.text ).strip( ';' )
        name = str( cells[1].font.text ).strip( ';* ' )

        code = self.codemap[ code ] if self.codemap.has_key( code ) else code

        current_date = self.meta.first_day

        # For each day ...
        for cell in cells[3:]:
            current = {}

            # All entries belong to font tags
            for f in cell.findAll( 'font' ):
                for key, val in f.attrs.items():
                    if key == 'color':
                        # Fonts with color attributes represent start/type data for a single event
                        e = WbcAllInOne.Event()
                        e.code = code
                        e.name = name
                        hour = int( f.text.strip() )
                        day = current_date.day
                        month = current_date.month
                        if hour >= 24:
                            hour = hour - 24
                            day = day + 1
                        if day >= 32:  # This works because WBC always starts in either the end of July or beginning of August
                            day = day - 31
                            month = month + 1
                        e.time = TZ.localize( current_date.replace( month=month, day=day, hour=hour ) )
                        e.type = self.colormap.get( val, None )
                        current[hour] = e

                    elif key == 'size':
                        # Fonts with size=-1 represent entry data for all events
                        text = str( f.text ).strip().split( '; ' )

                        if len( text ) == 1:
                            # If there's only one entry, it applies to all events
                            for e in current.values():
                                e.location = text[0]
                        else:
                            # For each entry ...
                            for chunk in text:
                                times, dummy, entry = chunk.partition( ':' )
                                if times == 'others':
                                    # Apply this location to all entries without locations
                                    for e in current.values():
                                        if not e.location:
                                            e.location = entry
                                else:
                                    # Apply this location to each listed hour
                                    for hour in times.split( ',' ):
                                        current[int( hour )].location = entry

            # Add all of this days events to the list
            events = events + current.values()

            # Move to the next date
            current_date = current_date + timedelta( days=1 )

        # Sort the list, then add it to the events map
        events.sort()
        self.events[code] = events

#----- WBC Yearbook Schedule -------------------------------------------------

class WbcYearbook( object ):
    """This class is used to parse schedule data from the annual yearbook pages"""

    # Basically there are two message streams to parse:
    #
    # 1) When events are happening
    # 2) Where events are happening
    #
    # When messages were originally framed as one day per table cell (<td>),
    # but with events that now stretch for 8 or more days, now some cells
    # may encompass as many as 5 days.  To complicate matters,
    # instead of indicating a date, images are used to indicate a day.  With
    # the convention stretching from 9 days from Saturday to the following Sunday,
    # there are two Saturdays and two Sundays, each represented by the same icon.

    PAGE_URL = "http://boardgamers.org/yearbkex/%spge.htm"
    INDEX_URL = "http://boardgamers.org/yearbkex%d/"

    # codemap = { 'MRA': 'MMA', }
    # codemap = { 'mma' : 'MRA' }
    codemap = { 'kot' : 'KOT' }

    # TODO: Preview codes for messages
    notes = {
        'CNS': "Can't match 30 minute rounds",
        'ELC': "Can't match 20 minute rounds",
        'LID': "Can't match 20 minute rounds",
        'PDT': "Can't match 30 minute drafts",
        'SLS': "Can't match 20 minute rounds",
        'SSB': "Can't match 30 minute drafts",
#       'WAW': "Can't handle midday Tuesday room switch",

        'ADV': "Preview shows 1 hour for SF and F, not 2 hours per spreadsheet and pocket schedule",
        'BAR': "Pocket schedule shows R1@8/8:9, R2@8/8:14, R3@8/8:19, SF@8/9:9, F@8/9:14",
        'KFE': "Pocket schedule shows R1@8/6:9, R2@8/6:16, SF@8/7:9, F@8/7:16",
        'KOT': "Can't handle 'to conclusion'",
        'LST': "Can't handle 'until conclusion'",
        'MED': "Preview shows heats taking place after SF/F",
        'PGF': "Can't handle 'to conclusion'",
        'RFG': "Should only have 1 demo",
        'STA': "Conestoga is misspelled as Coonestoga",


#       'EIS': 'Preview has split room name',
#       'KOH': 'Preview is missing time for last round',
#       'ROS': 'Preview has Wheatland misspelled as Wheatlamd',
#       'SQL': "Can't report on demos at the same time as events",
#       'T&T': 'Preview has H3 on Tuesday, not Thursday',
#       'TT2': 'Preview has demo at 21, instead of combined at 19',
    }

    events = {}

    valid = False

    class Event( object ):
        """Simple data object to collect information about an event occuring at a specific time."""

        code = None
        name = None
        type = None
        time = None
        location = None

        def __init__( self, code, name, etype, etime, location ):
            self.code = code
            self.name = name
            self.type = etype
            self.time = etime
            self.location = location

        def __cmp__( self, other ):
            return cmp( self.time, other.time )

        def __str__( self ):
            return '%s %s %s in %s at %s' % ( self.code, self.name, self.type, self.location, self.time )

    class Token( object ):
        """Simple data object for breaking descriptions into parseable tokens"""

        type = None
        value = None

        def __init__( self, type, value ):
            self.type = type
            self.value = value

    class Tourney( object ):
        """Class to organize events for a Yearbook tournament."""

        icon_meanings = {
            'stadium' : 'dummy',
            'demo' : 'Demo', 'demoweb' : 'Demo', 'demo_folder_transparent' : 'Demo',
            'jrwebicn': 'Junior', 'mulligan' : 'Mulligan', 'semi' : 'SF', 'final' : 'F',
            'heat1' : 'H1', 'heat2' : 'H2', 'heat3' : 'H3', 'heat4' : 'H4',
            'rd1' : 'R1', 'rd2' : 'R2', 'rd3' : 'R3', 'rd4' : 'R4', 'rd5' : 'R5', 'rd6' : 'R6',
            'sat2' : 'SAT', 'sun2' : 'SUN', 'mon2' : 'MON', 'tue2' : 'TUE', 'wed2' : 'WED', 'thu2' : 'THU', 'fri2' : 'FRI',
        }

        event_codes = [ 'H1', 'H2', 'H3', 'H4', 'R1', 'R2', 'R3', 'R4', 'R5', 'R6', 'Mulligan', 'Demo', 'After Action', 'Draft' ]

        days = { 'SAT' : 0, 'SUN' : 1, 'MON' : 2, 'TUE' : 3, 'WED' : 4, 'THU' : 5, 'FRI' : 6 }
        reverse = { 0 : 'SAT', 1 : 'SUN', 2 : 'MON', 3 : 'TUE', 4 : 'WED', 5 : 'THU', 6 : 'FRI' }

        # TODO: Preview codes to debug
        dump = [
            'AFK', 'AGE', 'B&O', 'BBS', 'BRI', 'BRS', 'BWD', 'CIS', 'CQT', 'FMR',
            'GBG', 'GBM', 'GSR', 'HWD', 'IVH', 'KRM', 'LHV', 'MFD', 'MMA', 'MOV',
            'PGD', 'POF', 'PRO', 'RFG', 'RRY', 'SFR', 'SMW', 'SPG', 'SPY', 'SSB',
            'STA', 'T&T', 'TRC', 'VSD', 'WAT', 'WPS', 'WSM',
        ]  # List of codes to dump for parser debugging

        default_room = None
        shift_room = None
        shift_time = None

        event_tokens = None
        room_tokens = None
        event_map = None
        events = None

        def __init__( self, code, name, page, first_day ):

            self.code = code
            self.name = name
            self.first_day = first_day

            # Find schedule / rows
            tables = page.findAll( 'table' )
            schedule = tables[2]
            rows = schedule.findAll( 'tr' )

            self.get_event_tokens( rows )
            self.get_room_tokens( rows )
            self.debug()
            self.create_event_map()
            self.create_events()
            self.events.sort()

        def get_event_tokens( self, rows ):
            """Parse 3rd row through next-to-last for event data"""

            self.event_tokens = []

            for row in rows[2:-1]:
                for td in row.findAll( 'td' ):
                    self.event_tokens.append( '|' )
                    for tag in td.descendants:
                        if isinstance( tag, NavigableString ):
                            token = self.clean_token( tag )
                            if token:
                                self.event_tokens.append( token )
                        elif isinstance( tag, Tag ) and tag.name in ( 'img' ):
                            event_type = self.parse_type( tag )
                            self.event_tokens.append( event_type )

        def get_room_tokens( self, rows ):
            """"parse last row for room data"""

            self.room_tokens = []

            for loc_data in rows[-1].findAll( 'center' ):
                for tag in loc_data.descendants:
                    if isinstance( tag, NavigableString ):
                        token = self.clean_token( tag )
                        if token:
                            self.room_tokens.append( token )
                    elif isinstance( tag, Tag ) and tag.name in ( 'img' ):
                        event_type = self.parse_type( tag )
                        self.room_tokens.append( event_type )

        def debug( self ):
            """Dump distilled data from webpage, before processing"""

            if not DEBUGGING or not self.code in self.dump:
                return

            LOGGER.warn( "%3s: rooms %s", self.code, self.flatten( self.room_tokens ) )
            LOGGER.warn( "     times %s", self.flatten( self.event_tokens ) )

            return

        @staticmethod
        def flatten( chunk_list ):
            chunks = []
            for tokens in chunk_list:
                if isinstance( tokens, list ):
                    time_tokens = '[' + ','.join( [ str( times.seconds / 3600 ) for times in tokens ] ) + ']'
                    chunks.append( time_tokens )
                else:
                    chunks.append( tokens )

            return '~'.join( chunks )

        def create_event_map( self ):
            """Parse room data to set rooms for events"""

            next_events = []

            self.event_map = {}

            remainder = ''
            for token in self.room_tokens:
                if remainder:
                    token = remainder + token
                    remainder = ''

                if token == 'dummy':
                    pass  # Skip the stadium icon
                elif isinstance( token, list ):
                    pass  # If we see a list, it will contain a timedelta that is really the table number on the Terrace -- ignore it.
                elif token in self.event_codes:
                    next_events.append( token )
                else:
                    # We're looking at a room name, possibly followed by a second room name and a time.
                    room, dummy, s = token.partition( '>' )
                    try:
                        r, dummy, t = s.partition( '@' )
                        if r and t:
                            d, dummy, t = t.partition( ':' )
                            self.shift_time = self.first_day + timedelta( days=self.days[d] ) + timedelta( hours=int( t ) )
                            self.shift_room = r
                    except:
                        self.shift_room = None

                    if room:
                        self.default_room = room if not self.default_room else self.default_room
                        for e in next_events:
                            self.event_map[e] = room
                    next_events = []

        def create_events( self ):
            """Parse event data to create events"""

            found_event_time = False
            etype = None
            weekend_offset = 0
            location = None
            last_day = None

            self.events = []

            for token in self.event_tokens:
                if token == '|':
                    # reset the frame
                    days = []
                    found_event_time = False
                elif isinstance( token, list ):
                    # event time -- add an event
                    found_event_time = True

                    # If there hasn't been a day specified in this frame, assume the day after the last one
                    if len( days ) == 0 and last_day:
                        days.append( self.reverse[( self.days[last_day] + 1 ) % 7] )

                    for day in days:
                        # If there are multiple days, only put demos on the last day of a block
                        if etype == 'Demo' and day != days[-1]:
                            continue

                        # Calculate which day of the convention (accounting for two weekends)
                        day_offset = self.days[ day ]
                        if day_offset < 2:
                            day_offset = day_offset + weekend_offset
                        d = self.first_day + timedelta( days=day_offset )

                        # For each time on this day, create a matching event
                        for t in token:
                            etime = d + t
                            if not location:
                                if self.shift_room and etime >= self.shift_time:
                                    location = self.shift_room
                                else:
                                    location = self.default_room

                            e = WbcYearbook.Event( self.code, self.name, etype, TZ.localize( etime ), location )
                            self.events.append( e )
                            location = None

                elif self.days.has_key( token ):
                    # event date
                    days.append( token )

                    # If this day isn't Saturday or Sunday, then the next Saturday or Sunday we see will be from the second weekend.
                    if self.days[token] > 1:
                        weekend_offset = 7

                    # Keep track of the last day seen
                    last_day = token

                else:
                    # event type
                    if token in [ 'SF', 'F', 'After Action' ]:
                        if found_event_time:
                            # if we found a time before these event types, then we've already created this event,
                            # and we need to patch it with the correct type (and location, if necessary).
                            self.events[-1].type = token
                            if self.event_map.has_key( token ):
                                self.events[-1].location = self.event_map[token]
                        else:
                            etype = token
                    elif token == 'PC':
                        self.events[-1].name = self.events[-1].name + ' PC'
                    else:
                        found_event_time = False
                        etype = token
                        location = self.event_map[etype] if self.event_map.has_key( etype ) else None

        @staticmethod
        def clean_token( ns ):
            """Remove extraneous characters and patterns from page text"""

            text = ns.strip()
            text = text.replace( '\n', '' )
            text = text.replace( u'\xa0', ' ' )
            text = text.replace( '#', 'Table ' )
            text = text.replace( ';', ',' )
            text = text.replace( ' ' * 8, ' ' ).replace( ' ' * 4, ' ' ).replace( ' ' * 2, ' ' )
            text = text.replace( ', Table', '' )
            text = text.replace( 'Draft:', 'Draft' ).replace( 'DRAFT:', 'Draft' )
            text = text.replace( 'After Action Briefing:', 'After Action' )
            text = text.replace( 'After Action Briefing', 'After Action' )
            text = text.replace( ' AFC', '' ).replace( ' NFC', '' ).replace( ' Super Bowl', '' )
            text = text.replace( ' to completion', '' ).replace( ' till completion', '' )
            text = text.replace( 'Grognard PC', 'PC' )
            text = text.replace( 'moves to ', '>' )
            text = text.replace( 'moving to ', '>' )
            text = text.replace( 'shifts to ', '>' )
            text = text.replace( 'switching to ', '>' )
            text = text.replace( ', >', '>' ).replace( ',>', '>' )
            text = text.replace( ' @ ', '@' )
            text = text.replace( '@ ', '@' )
            text = text.replace( '@We9', '@WED:9' )
            text = text.replace( '9-19', '9' )
            text = text.replace( '+', '' )
            text = text.rstrip( ', ' )
            try:
                items = text.split( ', ' )
                times = [ timedelta( hours=int( n ) ) for n in items if n > 8 ]
                return times
            except:
                return text

        def parse_type( self, tag ):
            """Remove extraneous characters from page type"""
            event_type = tag['src'].lower()
            event_type = event_type.split( '/' )[-1]
            event_type = event_type.split( '.' )[0]
            if self.icon_meanings.has_key( event_type ):
                event_type = self.icon_meanings[ event_type ]
            return event_type

    def __init__( self, metadata, options, event_names ):
        self.meta = metadata
        self.options = options

        self.names = event_names  # mapping of codes to event names
        self.codes = event_names.keys()
        self.codes.sort()

        self.yy = self.options.year % 100
        if self.options.year != self.meta.this_year:
            self.PAGE_URL = "http://boardgamers.org/yearbkex%d/%%spge.htm" % ( self.yy, )

        LOGGER.info( 'Loading Yearbook schedule' )
        index = parse_url( self.INDEX_URL % ( self.yy, ) )
        if not index:
            LOGGER.error( 'Unable to load Yearbook index' )

        for option in index.findAll( 'option' ):
            value = option['value']
            if value == 'none' or value == '' or value == 'jnrpge.htm':
                continue
            pagecode = value[0:3]
            self.load_yearbook_page( pagecode )

        self.valid = True

    def load_yearbook_page( self, pagecode ):
        """Load and parse the yearbook page for a single tournament
        
        The schedule table within the page is a table that has a variable number of rows:
        
            [0] Contains the date the page was last updated -- ignored.
            [1] Contains the token code and other image codes
            [2:-2] Contains the schedule data, mostly as images, in two columns
            [-1] Contains the location information.
            
        As is the case with all of the WBC web pages, the HTML is ugly and malformed.
        """

        LOGGER.debug( 'Loading yearbook for %s', pagecode )

        # Map page codes to event codes
        code = self.codemap[ pagecode ] if self.codemap.has_key( pagecode ) else pagecode.upper()

#         # Skip any codes whose pages we can't handle
#         if self.skip.has_key( code ):
#             LOGGER.warn( 'Skipping %s: %s -- %s', code, self.names[ code ], self.skip[ code ] )
#             return

        # Load page
        url = self.PAGE_URL % pagecode
        page = parse_url( url )
        if not page:
            LOGGER.error( "Unable to load %s for [%s:%s]", url, pagecode, code )
            return

        if self.names.has_key( code ):
            t = WbcYearbook.Tourney( code, self.names[ code ], page, self.meta.first_day )
            self.events[ code ] = t.events
        else:
            LOGGER.error( "No event name for code [%s]; not loading yearbook", code )

#----- Schedule Comparison ---------------------------------------------------

class ScheduleComparer( object ):
    """This class knows enough about the different schedule sources to compare events"""

    TEMPLATE = 'report-template.html'

    def __init__( self, metadata, options, s, a, y=None ):
        self.options = options
        self.meta = metadata
        self.schedule = s
        self.allinone = a
        self.yearbook = y
        self.parser = None

    def verify_event_calendars( self ):
        """Compare the collections of events from both the calendars and the schedule"""

        LOGGER.info( 'Verifying event calendars against All-in-One schedule' )

        schedule_key_set = set( self.schedule.current_tourneys )

        if self.allinone.valid:
            allinone_key_set = set( self.allinone.events.keys() )
            allinone_extras = allinone_key_set - schedule_key_set
            allinone_omited = schedule_key_set - allinone_key_set
        else:
            allinone_extras = set()
            allinone_omited = set()

        if self.yearbook.valid:
            yearbook_key_set = set( self.yearbook.events.keys() )
            yearbook_extras = yearbook_key_set - schedule_key_set
            yearbook_omited = schedule_key_set - yearbook_key_set
        else:
            yearbook_extras = set()
            yearbook_omited = set()

        add_space = False
        if len( allinone_extras ):
            LOGGER.error( 'Extra events present in All-in-One: %s', allinone_extras )
            add_space = True
        if len( allinone_omited ):
            LOGGER.error( 'Events omitted in All-in-One: %s', allinone_omited )
            add_space = True
        if len( yearbook_extras ):
            LOGGER.error( 'Extra events present in Yearbook: %s', yearbook_extras )
            add_space = True
        if len( yearbook_omited ):
            LOGGER.error( 'Events omitted in Yearbook: %s', yearbook_omited )
            add_space = True
        if add_space:
            LOGGER.error( '' )

        code_set = schedule_key_set
        if self.allinone.valid:
            code_set = code_set & allinone_key_set
        if self.yearbook.valid:
            code_set = code_set & yearbook_key_set

        codes = list( code_set )
        codes.sort()

        self.initialize_discrepancies_report()
        for code in codes:
            self.report_discrepancies( code )
        self.write_discrepancies_report()

    def initialize_discrepancies_report( self ):
        """Initial discrepancies report from template"""

        with open( self.TEMPLATE, "r" ) as f:
            template = f.read()

        self.parser = BeautifulSoup( template )

        # Locate insertion points
        title = self.parser.find( 'title' )
        header = self.parser.find( 'div', { 'id' : 'header' } )
        footer = self.parser.find( 'div', { 'id' : 'footer' } )

        # Page title
        if self.options.fullreport:
            text = "WBC %s Schedule Details" % self.options.year
        else:
            text = "WBC %s Schedule Discrepancies" % self.options.year

        title.insert( 0, self.parser.new_string( text ) )
        header.h1.insert( 0, self.parser.new_string( text ) )
        footer.p.insert( 0, self.parser.new_string( "Updated on %s" % self.schedule.processed.strftime( "%A, %d %B %Y %H:%M %Z" ) ) )

    def report_discrepancies( self, code ):
        """Format the discrepancies for a given tournament"""

        # Find all of the matching events from each schedule
        ai1_events = self.allinone.events[ code ] if self.allinone.valid else []
        ybk_events = self.yearbook.events[ code ] if self.yearbook.valid else []
        ybk_events = [ e for e in ybk_events if e.type != 'Junior' ]
        cal_events = self.schedule.calendars[code].subcomponents

        # Find all of the unique times for any events
        ai1_timemap = dict( [ ( e.time.astimezone( TZ ), e ) for e in ai1_events ] )
        ybk_timemap = dict( [ ( e.time.astimezone( TZ ), e ) for e in ybk_events ] )
        cal_timemap = dict( [ ( e['dtstart'].dt.astimezone( TZ ), e ) for e in cal_events ] )
        time_set = set( ai1_timemap.keys() ) | set( ybk_timemap.keys() ) | set( cal_timemap.keys() )
        time_list = list( time_set )
        time_list.sort()

        label = self.meta.names[code]

        discrepancies = False

        rows = []
        self.create_discrepancy_header( rows, code )

        # For each date/time combination, compare all of the events at that time
        for starting_time in time_list:
            # Start with empty cells
            details = [( None, ), ( None, ), ( None, ), ]

            # Fill in the All-in-One event, if present
            if ai1_timemap.has_key( starting_time ):
                e = ai1_timemap[ starting_time ]
                location = 'Terrace' if e.location == 'Pt' else e.location
                details[0] = ( location, e.type )

            # Fiil in the Preview event, if present
            if ybk_timemap.has_key( starting_time ):
                e = ybk_timemap[ starting_time ]
                location = 'Terrace' if e.location and e.location.startswith( 'Terr' ) else e.location
                details[1] = ( location, e.type )

            # Fill in the spreadsheet event, if present
            if cal_timemap.has_key( starting_time ):
                e = cal_timemap[ starting_time ]
                location = 'Terrace' if e['location'].startswith( 'Terr' ) else e['location']
                summary = unicode( e['summary'] )
                try:
                    ulab = codecs.decode( label, 'utf-8' )
                    summary = summary[len( ulab ) + 1:] if summary.startswith( ulab ) else summary
                except Exception as x:
                    LOGGER.error( u'Could not handle %s', summary )
                    LOGGER.exception( x )
                seconds = e['duration'].dt.seconds
                hours = int( seconds / 3600 )
                minutes = int ( ( seconds - 3600 * hours ) / 60 )
                duration = "%d:%02d" % ( hours, minutes )
                details[2] = ( location, summary, duration )

            result = self.add_discrepancy_row( rows, starting_time, details )
            discrepancies = discrepancies or result

        # If we have notes, add them
        if self.yearbook.valid and self.yearbook.notes.has_key( code ) :
            discrepancies = True
            tr = self.parser.new_tag( 'tr' )
            td = self.parser.new_tag( 'td' )
            td['colspan'] = 8
            td['class'] = 'note'
            td.insert( 0, self.parser.new_string( self.yearbook.notes[ code ] ) )
            tr.insert( len( tr ), td )
            rows.append( tr )

        # Set the correct row span for this event
        rows[0].next['rowspan'] = len( rows )

        # If there were discrepancies, then add the rows to the report
        if discrepancies or self.options.fullreport:
            self.add_discrepancies_to_report( rows )

    def create_discrepancy_header( self, rows, code ):
        """Create the first row of the discrepancies table (code, name, headers)"""

        label = self.meta.names[code]

        tr = self.parser.new_tag( 'tr' )

        th = self.parser.new_tag( 'th' )
        th['class'] = 'eventcode'
        th.insert( 0, self.parser.new_string( code ) )
        tr.insert( len( tr ), th )

        th = self.parser.new_tag( 'th' )
        th['class'] = 'eventname'
        th.insert( 0, self.parser.new_string( label ) )
        tr.insert( len( tr ), th )

        if self.allinone.valid:
            th = self.parser.new_tag( 'th' )
            th['colspan'] = 2
            th.insert( 0, self.parser.new_string( 'All-in-One' ) )
            tr.insert( len( tr ), th )

        if self.yearbook.valid:
            a = self.parser.new_tag( 'a' )
            a['href'] = self.yearbook.PAGE_URL % code.lower()
            a.insert( 0, self.parser.new_string( 'Event Preview' ) )
            th = self.parser.new_tag( 'th' )
            th['colspan'] = 2
            th.insert( 0, a )
            tr.insert( len( tr ), th )

        th = self.parser.new_tag( 'th' )
        th['colspan'] = 3
        th.insert( 0, self.parser.new_string( 'Spreadsheet' ) )
        tr.insert( len( tr ), th )

        rows.append( tr )

    def add_discrepancy_row( self, rows, starting_time, details ):
        """Format a discrepancy row for this time"""

        if not self.yearbook.valid:
            del details[1]

        if not self.allinone.valid:
            del details[0]

        # Calculate which calendars are different than the others
        if len( details ) == 1:
            differences = set()
        elif len( details ) == 2:
            if details[0][0] == details[1][0]:
                differences = set()
            else:
                differences = set( [0, 1 ] )
        elif details[0][0] == details[1][0] and details[1][0] == details[2][0]:
            differences = set()
        elif details[0][0] == details[1][0]:
            differences = set( [2] )
        elif details[0][0] == details[2][0]:
            differences = set( [1] )
        elif details[1][0] == details[2][0]:
            differences = set( [0] )
        else:
            differences = set( [0, 1, 2] )

        # Create a new row
        tr = self.parser.new_tag( 'tr' )

        # Add the starting time for this row
        td = self.parser.new_tag( 'td' )
        td.insert( 0, self.parser.new_string( starting_time.strftime( '%a %m-%d %H:%M' ) ) )
        tr.insert( len( tr ), td )

        # For each detailed event, create appropriately marked cells
        for i in range( len( details ) ):
            if details[i][0] == None:
                td = self.parser.new_tag( 'td' )
                td['colspan'] = 2 if i < len( details ) - 1 else 3
                if i in differences:
                    td['class'] = 'diff'
                tr.insert( len( tr ), td )
            else:
                for j in range( len( details[i] ) ):
                    td = self.parser.new_tag( 'td' )
                    if i in differences:
                        td['class'] = 'diff'
                    value = '' if details[i][j] == None else details[i][j]
                    td.insert( 0, self.parser.new_string( value ) )
                    tr.insert( len( tr ), td )

        rows.append( tr )

        return len( differences ) > 0

    def add_discrepancies_to_report( self, rows ):
        """Add these rows to the discrepancies report"""

        table = self.parser.find( 'div', {'id':'details'} ).table

        if len( table ):
            # Add a blank cell for padding
            tr = self.parser.new_tag( 'tr' )
            td = self.parser.new_tag( 'td' )
            td['colspan'] = 9
            tr.insert( len( tr ), td )
            table.insert( len( table ), tr )

        for row in rows:
            table.insert( len( table ), row )

    def write_discrepancies_report( self ):
        """Write the discrepancies report, in a nice pretty format"""

        path = os.path.join( self.schedule.options.output, "report.html" )
        with open( path, "w" ) as f:
            f.write( self.parser.prettify() )

    def check_schedules_against_each_other( self, code ):
        """Compare all of the schedules against each other, logging the differences"""

        if not self.allinone.valid or not self.yearbook.valid:
            return

        ai1_events = self.allinone.events[ code ]
        ybk_events = self.yearbook.events[ code ]
        cal_events = self.schedule.calendars[code].subcomponents

        # Remove calendar events that aren't tracked on all schedules
        cal_events = [ e for e in cal_events if not e['summary'].find( ' Jr' ) >= 0 ]
        cal_events = [ e for e in cal_events if not e['summary'].endswith( 'After Action' ) ]
        cal_events = [ e for e in cal_events if not e['summary'].endswith( 'Draft' ) ]
        cal_events = [ e for e in cal_events if e['dtstart'].dt.minute == 0 ]
        ybk_events = [ e for e in ybk_events if e.type != 'Junior' ]
        ybk_events = [ e for e in ybk_events if e.type != 'After Action' ]

        # Create lists of the comparable information from both sets of events
        ai1_comparison = [ self.ai1_date_loc( x ) for x in ai1_events ]
        ybk_comparison = [ self.ybk_date_loc( x ) for x in ybk_events ]
        cal_comparison = [ self.cal_date_loc( x ) for x in cal_events ]

        # Compare the lists looking for discrepancies
        ac_changes = set( ai1_comparison ) - set( cal_comparison )
        ca_changes = set( cal_comparison ) - set( ai1_comparison )
        ay_changes = set( ai1_comparison ) - set( ybk_comparison )
        ya_changes = set( ybk_comparison ) - set( ai1_comparison )
        yc_changes = set( ybk_comparison ) - set( cal_comparison )
        cy_changes = set( cal_comparison ) - set( ybk_comparison )

        changes = len( ac_changes ) or len( ca_changes ) or len( ay_changes ) or len( ya_changes ) or len( yc_changes ) or len( cy_changes )

        if changes:
            LOGGER.error( '%s: All-in-One_____________________ Yearbook_______________________ Spreadsheet____________________ __Length Name________________', code )
            cal_other = [ '%8s %s' % ( x['duration'].dt, x['summary'] ) for x in cal_events ]
            for ai1, ybk, cal, other in izip_longest( ai1_comparison, ybk_comparison, cal_comparison, cal_other, fillvalue='' ):
                mark = ' ' if ai1 == ybk and ybk == cal else '*'
                LOGGER.error( '%4s %-31s %-31s %-31s %s', mark, ai1, ybk, cal, other )
            LOGGER.error( '' )

    def check_schedule_against_allinone( self, code ):
        """Compare the generated calendar for an event against the all-in-one schedule 
        for that event.  If there are differences, log them."""

        if not self.allinone.valid:
            return

        ai1_events = self.allinone.events[ code ]
        cal_events = self.schedule.calendars[code].subcomponents

        # Remove calendar events that aren't tracked on the all-in-one schedule
        cal_events = [ e for e in cal_events if not e['summary'].find( ' Jr' ) >= 0 ]
        cal_events = [ e for e in cal_events if not e['summary'].endswith( 'After Action' ) ]
        cal_events = [ e for e in cal_events if not e['summary'].endswith( 'Draft' ) ]
        cal_events = [ e for e in cal_events if e['dtstart'].dt.minute == 0 ]

        # Create lists of the comparable information from both sets of events
        ai1_comparison = [ self.ai1_date_loc( x ) for x in ai1_events ]
        cal_comparison = [ self.cal_date_loc( x ) for x in cal_events ]

        # Compare the lists looking for discrepancies
        ai1_extra = set( ai1_comparison ) - set( cal_comparison )
        cal_extra = set( cal_comparison ) - set( ai1_comparison )

        # If there are any discrepancies, log them
        if len( ai1_extra ) or len( cal_extra ):
            LOGGER.error( '%s: All-in-One_____________________ Spreadsheet____________________ __Length Name________________', code )
            cal_other = [ '%8s %s' % ( x['duration'].dt, x['summary'] ) for x in cal_events ]
            for tab, cal, other in izip_longest( ai1_comparison, cal_comparison, cal_other, fillvalue='' ):
                mark = ' ' if tab == cal else '*'
                LOGGER.error( '%4s %-31s %-31s %s', mark, tab, cal, other )
            LOGGER.error( '' )

    def check_schedule_against_yearbook( self, code ):
        """Compare the generated calendar for an event against the yearbook schedule 
        for that event.  If there are differences, log them."""

        if not self.yearbook.valid:
            return

        ybk_events = self.yearbook.events[ code ]
        cal_events = self.schedule.calendars[code].subcomponents

        # Remove calendar events that aren't tracked on the yearbook schedule
#       cal_events = [ e for e in cal_events if not e['summary'].find( ' Jr' ) >= 0 ]
#       cal_events = [ e for e in cal_events if not e['summary'].endswith( 'After Action' ) ]
#       cal_events = [ e for e in cal_events if not e['summary'].endswith( 'Draft' ) ]
        cal_events = [ e for e in cal_events if e['dtstart'].dt.minute == 0 ]
        ybk_events = [ e for e in ybk_events if e.type != 'Junior' ]

        # Create lists of the comparable information from both sets of events
        ybk_comparison = [ self.ybk_date_loc( x ) for x in ybk_events ]
        cal_comparison = [ self.cal_date_loc( x ) for x in cal_events ]

        # Compare the lists looking for discrepancies
        ybk_extra = set( ybk_comparison ) - set( cal_comparison )
        cal_extra = set( cal_comparison ) - set( ybk_comparison )

        # If there are any discrepancies, log them
        if len( ybk_extra ) or len( cal_extra ):
            LOGGER.error( '%s: Yearbook_______________________ Spreadsheet____________________ __Length Name________________', code )
            cal_other = [ '%8s %s' % ( x['duration'].dt, x['summary'] ) for x in cal_events ]
            for tab, cal, other in izip_longest( ybk_comparison, cal_comparison, cal_other, fillvalue='' ):
                mark = ' ' if tab == cal else '*'
                LOGGER.error( '%4s %-31s %-31s %s', mark, tab, cal, other )
            LOGGER.error( '' )

    def check_allinone_against_yearbook( self, code ):
        """Check the All-in-One schedule against the Yearbook schedule"""

        if not self.allinone.valid or not self.yearbook.valid:
            return

        ai1_events = self.allinone.events[ code ]
        ybk_events = self.yearbook.events[ code ]

        # Remove yearbook events that aren't tracked on the all-in-one schedule

        # Create lists of the comparable information from both sets of events
        ai1_comparison = [ self.ai1_date_loc( x ) for x in ai1_events ]
        ybk_comparison = [ self.ybk_date_loc( x ) for x in ybk_events ]

        # Compare the lists looking for discrepancies
        ai1_extra = set( ai1_comparison ) - set( ybk_comparison )
        ybk_extra = set( ybk_comparison ) - set( ai1_comparison )

        # If there are any discrepancies, log them
        if len( ybk_extra ) or len( ai1_extra ):
            LOGGER.error( '%s: All-in-One_____________________ Yearbook_______________________', code )
            for tab, cal in izip_longest( ai1_comparison, ybk_comparison, fillvalue='' ):
                mark = ' ' if tab == cal else '*'
                LOGGER.error( '%4s %-31s %-31s', mark, tab, cal )
            LOGGER.error( '' )

    @staticmethod
    def ai1_date_loc( ev ):
        """Generate a summary of an all-in-one event for comparer purposes"""

        start_time = ev.time.astimezone( TZ )
        location = ev.location
        location = 'Terrace' if location == 'Pt' else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M' ), location )

    @staticmethod
    def ybk_date_loc( ev ):
        """Generate a summary of an yearbook event for comparer purposes"""

        start_time = ev.time.astimezone( TZ )
        location = ev.location
        location = 'Terrace' if location.startswith( 'Terr' ) else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M' ), location )

    @staticmethod
    def cal_date_loc( sc ):
        """Generate a summary of a calendar event for comparer purposes"""

        start_time = sc['dtstart'].dt.astimezone( TZ )
        location = sc['location']
        location = 'Terrace' if location.startswith( 'Terrace' ) else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M' ), location )

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

    # Parse the WBC Yearbook
    names = dict( [( key_code, val_name ) for key_code, val_name in meta.names.items() if key_code in wbc_schedule.current_tourneys ] )
    wbc_yearbook = WbcYearbook( meta, opts, names )

    # Compare the event calendars with the WBC All-in-One schedule and the yearbook
    comparer = ScheduleComparer( meta, opts, wbc_schedule, wbc_allinone, wbc_yearbook )
    comparer.verify_event_calendars()

    LOGGER.warn( "Done." )
