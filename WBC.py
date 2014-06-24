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

from bs4 import BeautifulSoup, Tag, NavigableString, Comment
from icalendar import Calendar, Event

import pytz
import xlrd

logging.basicConfig( level=logging.INFO )
LOGGER = logging.getLogger( 'WBC' )

DEBUGGING = True
TRAPPING = False

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
        if DEBUGGING and self.name.find( 'Titan' ) >= 0:
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
                        raise ValueError( 'Unable to parse (%s) as a time' % self.time )

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
        if DEBUGGING and event.code in [ 'TTN' ]:
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
        # FIXME: This is wrong for BWD, which starts at 10am on the PC days, not 9am
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

        spreadsheet_file = os.path.join( self.options.output, "schedule.csv" )
        with codecs.open( spreadsheet_file, "w", 'utf-8' ) as f:
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
            span.insert( 0, parser.new_string( key + ': ' ) )
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
            a.insert( 0, parser.new_string( "%s" % label ) )
            td.insert( len( td ), a )

            td.insert( len( td ), a )
        else:
            td.insert( len( td ) , parser.new_string( ' ' ) )
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
        a.insert( 0, parser.new_string( "%s" % label ) )
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

#----- Token -----------------------------------------------------------------

class Token( object ):
    """Simple data object for breaking descriptions into parseable tokens"""

    INITIALIZED = False

    type = None
    label = None
    value = None

    def __init__( self, t, l=None, v=None ):
        self.type = t
        self.label = l
        self.value = v

    def __str__( self ):
        return str( self.label ) if self.label else self.type

    def __repr__( self ):
        return self.__str__()

    def __eq__( self, other ):
        return self.type == other.type and self.label == other.label and self.value == other.value

    def __ne__( self, other ):
        return not self.__eq__( other )

    @classmethod
    def initialize( cls ):
        if cls.INITIALIZED: return

        cls.START = Token( 'Symbol', '|' )
        cls.AT = Token( 'Symbol', '@' )
        cls.SHIFT = Token( 'Symbol', '>' )
        cls.PLUS = Token( 'Symbol', '+' )
        cls.DASH = Token( 'Symbol', '-' )
        cls.CONTINUOUS = Token( 'Symbol', '...' )

        cls.DAYS = {}
        cls.DAYS[ 'SAT' ] = cls.DAYS[ 0 ] = Token( 'Day', 'SAT', 0 )
        cls.DAYS[ 'SUN' ] = cls.DAYS[ 1 ] = Token( 'Day', 'SUN', 1 )
        cls.DAYS[ 'MON' ] = cls.DAYS[ 2 ] = Token( 'Day', 'MON', 2 )
        cls.DAYS[ 'TUE' ] = cls.DAYS[ 3 ] = Token( 'Day', 'TUE', 3 )
        cls.DAYS[ 'WED' ] = cls.DAYS[ 4 ] = Token( 'Day', 'WED', 4 )
        cls.DAYS[ 'THU' ] = cls.DAYS[ 5 ] = Token( 'Day', 'THU', 5 )
        cls.DAYS[ 'FRI' ] = cls.DAYS[ 6 ] = Token( 'Day', 'FRI', 6 )

        cls.LOOKUP = {}
        cls.LOOKUP[ 'Ballroom A' ] = Token( 'Room', 'Ballroom A' )
        cls.LOOKUP[ 'Ballroom B' ] = Token( 'Room', 'Ballroom B' )
        cls.LOOKUP[ 'Ballroom AB' ] = Token( 'Room', 'Ballroom AB' )
        cls.LOOKUP[ 'Ballroom' ] = cls.LOOKUP[ 'Ballroom AB' ]

        cls.LOOKUP[ 'Conestoga 1' ] = Token( 'Room', 'Conestoga 1' )
        cls.LOOKUP[ 'Conestoga 2' ] = Token( 'Room', 'Conestoga 2' )
        cls.LOOKUP[ 'Conestoga 3' ] = Token( 'Room', 'Conestoga 3' )
        cls.LOOKUP[ 'Coonestoga 3'] = cls.LOOKUP[ 'Conestoga 3' ]

        cls.LOOKUP[ 'Cornwall' ] = Token( 'Room', 'Cornwall' )
        cls.LOOKUP[ 'Heritage' ] = Token( 'Room', 'Heritage' )
        cls.LOOKUP[ 'Hopewell' ] = Token( 'Room', 'Hopewell' )
        cls.LOOKUP[ 'Kinderhook' ] = Token( 'Room', 'Kinderhook' )
        cls.LOOKUP[ 'Lampeter' ] = Token( 'Room', 'Lampeter' )
        cls.LOOKUP[ 'Laurel Grove' ] = Token( 'Room', 'Laurel Grove' )
        cls.LOOKUP[ 'Limerock' ] = Token( 'Room', 'Limerock' )
        cls.LOOKUP[ 'Marietta' ] = Token( 'Room', 'Marietta' )
        cls.LOOKUP[ 'New Holland' ] = Token( 'Room', 'New Holland' )
        cls.LOOKUP[ 'Paradise' ] = Token( 'Room', 'Paradise' )
        cls.LOOKUP[ 'Showroom' ] = Token( 'Room', 'Showroom' )
        cls.LOOKUP[ 'Strasburg' ] = Token( 'Room', 'Strasburg' )
        cls.LOOKUP[ 'Wheatland' ] = Token( 'Room', 'Wheatland' )

        cls.LOOKUP[ 'Terrace 1' ] = Token( 'Room', 'Terrace 1' )
        cls.LOOKUP[ 'Terrace 2' ] = Token( 'Room', 'Terrace 2' )
        cls.LOOKUP[ 'Terrace 3' ] = Token( 'Room', 'Terrace 3' )
        cls.LOOKUP[ 'Terrace 4' ] = Token( 'Room', 'Terrace 4' )
        cls.LOOKUP[ 'Terrace 5' ] = Token( 'Room', 'Terrace 5' )
        cls.LOOKUP[ 'Terrace 6' ] = Token( 'Room', 'Terrace 6' )
        cls.LOOKUP[ 'Terrace 7' ] = Token( 'Room', 'Terrace 7' )

        cls.LOOKUP[ 'Vista C' ] = Token( 'Room', 'Vista C' )
        cls.LOOKUP[ 'Vista D' ] = Token( 'Room', 'Vista D' )
        cls.LOOKUP[ 'Vista CD' ] = Token( 'Room', 'Vista CD' )
        cls.LOOKUP[ 'Vista' ] = cls.LOOKUP[ 'Vista CD' ]

        cls.LOOKUP[ 'H1' ] = Token( 'Event', 'H1' )
        cls.LOOKUP[ 'H2' ] = Token( 'Event', 'H2' )
        cls.LOOKUP[ 'H3' ] = Token( 'Event', 'H3' )
        cls.LOOKUP[ 'H4' ] = Token( 'Event', 'H4' )
        cls.LOOKUP[ 'R1' ] = Token( 'Event', 'R1' )
        cls.LOOKUP[ 'R2' ] = Token( 'Event', 'R2' )
        cls.LOOKUP[ 'R3' ] = Token( 'Event', 'R3' )
        cls.LOOKUP[ 'R4' ] = Token( 'Event', 'R4' )
        cls.LOOKUP[ 'R5' ] = Token( 'Event', 'R5' )
        cls.LOOKUP[ 'R6' ] = Token( 'Event', 'R6' )
        cls.LOOKUP[ 'SF' ] = Token( 'Event', 'SF' )
        cls.LOOKUP[ 'F' ] = Token( 'Event', 'F' )
        cls.LOOKUP[ 'Demo' ] = Token( 'Event', 'Demo' )
        cls.LOOKUP[ 'Junior' ] = Token( 'Event', 'Junior' )
        cls.LOOKUP[ 'Mulligan' ] = cls.LOOKUP[ 'mulligan' ] = Token( 'Event', 'Mulligan' )
        cls.LOOKUP[ 'After Action' ] = cls.LOOKUP[ 'After Action Briefing' ] = Token( 'Event', 'After Action' )
        cls.LOOKUP[ 'Draft' ] = cls.LOOKUP[ 'DRAFT' ] = Token( 'Event', 'Draft' )

        cls.LOOKUP[ 'PC' ] = cls.LOOKUP[ 'Grognard PC' ] = Token( 'Qualifier', 'PC' )
        cls.LOOKUP[ 'AFC' ] = Token( 'Qualifier', 'AFC' )
        cls.LOOKUP[ 'NFC' ] = Token( 'Qualifier', 'NFC' )
        cls.LOOKUP[ 'Super Bowl' ] = Token( 'Qualifier', 'Super Bowl' )

        cls.LOOKUP[ 'to completion' ] = Token.CONTINUOUS
        cls.LOOKUP[ 'till completion' ] = Token.CONTINUOUS
        cls.LOOKUP[ 'until completion' ] = Token.CONTINUOUS
        cls.LOOKUP[ 'until conclusion' ] = Token.CONTINUOUS
        cls.LOOKUP[ 'to conclusion' ] = Token.CONTINUOUS

        cls.LOOKUP[ 'moves to' ] = Token.SHIFT
        cls.LOOKUP[ 'moving to' ] = Token.SHIFT
        cls.LOOKUP[ 'shifts to' ] = Token.SHIFT
        cls.LOOKUP[ 'switches to' ] = Token.SHIFT
        cls.LOOKUP[ 'switching to' ] = Token.SHIFT
        cls.LOOKUP[ 'after drafts in' ] = Token.SHIFT

        cls.PATTERN = '|'.join( sorted( Token.LOOKUP.keys(), reverse=True ) )

        cls.LOOKUP[ '@' ] = Token.AT
        cls.LOOKUP[ '+' ] = Token.PLUS
        cls.LOOKUP[ '-' ] = Token.DASH

        cls.PATTERN += '|[@+-]'

        cls.ICONS = {
            'semi' : Token( 'Award', 'SF' ),
            'final' : Token( 'Award', 'F' ),
            'heat1' : cls.LOOKUP['H1'],
            'heat2' : cls.LOOKUP['H2'],
            'heat3' : cls.LOOKUP['H3'],
            'heat4' : cls.LOOKUP['H4'],
            'rd1' : cls.LOOKUP['R1'],
            'rd2' : cls.LOOKUP['R2'],
            'rd3' : cls.LOOKUP['R3'],
            'rd4' : cls.LOOKUP['R4'],
            'rd5' : cls.LOOKUP['R5'],
            'rd6' : cls.LOOKUP['R6'],
            'demo' : cls.LOOKUP['Demo'],
            'demoweb' : cls.LOOKUP['Demo'],
            'demo_folder_transparent' : cls.LOOKUP['Demo'],
            'jrwebicn' : cls.LOOKUP[ 'Junior'],
            'mulligan' : cls.LOOKUP[ 'Mulligan' ],
            'sat' : cls.DAYS['SAT'],
            'sun' : cls.DAYS['SUN'],
            'mon' : cls.DAYS['MON'],
            'tue' : cls.DAYS['TUE'],
            'wed' : cls.DAYS['WED'],
            'thu' : cls.DAYS['THU'],
            'fri' : cls.DAYS['FRI'],
            'sat2' : cls.DAYS['SAT'],
            'sun2' : cls.DAYS['SUN'],
            'mon2' : cls.DAYS['MON'],
            'tue2' : cls.DAYS['TUE'],
            'wed2' : cls.DAYS['WED'],
            'thu2' : cls.DAYS['THU'],
            'fri2' : cls.DAYS['FRI'],
        }

    @staticmethod
    def dump_list( tokenlist ):
        return u'~'.join( [ unicode( x ) for x in tokenlist  ] )

    @classmethod
    def tokenize( cls, tag ):
        tokens = []
        buffer = u''

        if TRAPPING:
            pass

        for tag in tag.descendants:
            if isinstance( tag, Comment ):
                pass  # Always ignore comments
            elif isinstance( tag, NavigableString ):
                buffer += u' ' + unicode( tag )
            elif isinstance( tag, Tag ) and tag.name in ( 'img' ):
                tokens += cls.tokenize_text( buffer )
                buffer = u''
                tokens += cls.tokenize_icon( tag )
            else:
                pass  # ignore other tags, for now
                # LOGGER.warn( 'Ignored <%s>', tag.name )

        if buffer:
            tokens += cls.tokenize_text( buffer )

        return tokens

    @classmethod
    def tokenize_icon( cls, tag ):
        cls.initialize()

        tokens = []

        try:
            name = tag['src'].lower()
            name = name.split( '/' )[-1]
            name = name.split( '.' )[0]
        except:
            LOGGER.error( "%s didn't have a 'src' attribute", tag )
            return tokens

        if cls.ICONS.has_key( name ):
            tokens.append( cls.ICONS[ name ] )
        elif name in [ 'stadium' ]:
            pass
        else:
            LOGGER.warn( 'Ignored icon [%s]', name )

        return tokens

    @classmethod
    def tokenize_text( cls, text ):
        cls.initialize()

        data = text

        junk = u''
        tokens = []

        # Cleanup crappy data
        data = data.replace( u'\xa0', u' ' )
        data = data.replace( u'\n', u' ' )

        data = data.strip()
        data = data.replace( u' ' * 11, u' ' ).replace( u' ' * 7, u' ' ).replace( u' ' * 5, u' ' )
        data = data.replace( u' ' * 3, u' ' ).replace( u' ' * 2, u' ' )
        data = data.replace( u' ' * 2, u' ' ).replace( u' ' * 2, u' ' )
        data = data.strip()

        hdata = data.encode( 'unicode_escape' )

        while len( data ):

            # Ignore commas and semi-colons
            if data[0] in u',;:':
                data = data[1:]
            else:
                # Match Room names, event names, phrases, symbols
                m = re.match( cls.PATTERN, data )
                if m:
                    n = m.group()
                    tokens.append( cls.LOOKUP[ n ] )
                    data = data[len( n ):]
                else:
                    # Match numbers
                    m = re.match( "\d+", data )
                    if m:
                        n = m.group()
                        tokens.append( Token( 'Time', int( n ), timedelta( hours=int( n ) ) ) )
                        data = data[len( n ):]
                    else:
                        junk += data[0];
                        data = data[1:]

            data = data.strip()

        if junk:
            hjunk = junk.encode( 'unicode_escape' )
            LOGGER.warn( 'Skipped [%s] in [%s]', hjunk, hdata )

        return tokens


class Parser( object ):

    tokens = []

    def __init__( self, tokens ):
        self.tokens = list( tokens )
        self.last_match = None

    @property
    def count( self ):
        return len( self.tokens )

    def have( self, pos, *tokens ):
        tlen = len ( tokens )
        if len( self.tokens ) < pos + tlen:
            return False

        for i in range( tlen ):
            t = tokens[i]
            x = self.tokens[pos + i]
            if isinstance( t, Token ):
                if x != t:
                    return False
                else:
                    continue
            elif self.tokens[pos + i].type != t:
                return False

        return True

    def have_start( self, pos=0 ):
        return self.have( pos, Token.START )

    def have_day( self, pos=0 ):
        return self.have( pos, 'Day' )

    def recover( self ):
        p = 0
        while p < self.count and not self.have_start( p ) and not self.have_day( p ):
            p += 1

        if p < self.count:
            skipped = self.tokens[0:p]
            del self.tokens[0:p]
            LOGGER.warn( 'Recovered to %s by skipping %s', self.tokens[0], skipped )
            pass
        else:
            LOGGER.warn( 'Discarded remaining tokens: %s', self.tokens )
            self.tokens = []

    def match_initialize( self ):
        self.last_tokens = None
        self.last_match = None
        self.last_name = None
        self.last_start = None
        self.last_end = None
        self.last_room = None
        self.time_list = None
        self.last_continuous = False
        self.default_room = None
        self.shift_room = None
        self.shift_day = None
        self.shift_time = None
        self.event_list = None
        self.last_actual = None

    def match( self, token, pos=0 ):
        if self.have( pos, token ):
            del self.tokens[pos]
            LOGGER.debug( 'Matched %s', token )
            return 1
        else:
            return 0

    def match_day( self, pos=0 ):
        self.match_initialize()

        if not self.have( pos, 'Day' ):
            return 0

        self.last_day = self.tokens[pos]

        del self.tokens[pos]

        LOGGER.debug( 'Matched %s', self.last_day )
        return 1

    def match_single_event_time( self, pos=0, awards_are_events=False ):
        self.match_initialize()

        if self.have( pos, 'Event', 'Time' ):
            self.last_name = self.tokens[pos].label
            self.last_actual = self.tokens[pos].label
            self.last_start = self.tokens[pos + 1].value
            l = 2
        elif self.have( pos, 'Time', 'Event' ):
            self.last_start = self.tokens[pos].value
            self.last_name = self.tokens[pos + 1].label
            self.last_actual = self.tokens[pos + 1].label
            l = 2
        elif awards_are_events and self.have( pos, 'Award', 'Time' ):
            self.last_name = self.tokens[pos].label
            self.last_actual = self.tokens[pos].label
            self.last_start = self.tokens[pos + 1].value
            l = 2
        elif self.have( pos, 'Event', 'Award', 'Time' ):
            self.last_name = self.tokens[pos + 1].label
            self.last_actual = self.tokens[pos].label
            self.last_start = self.tokens[pos + 2].value
            l = 3
        else:
            return 0

        if self.have( pos + l, Token.DASH, 'Time' ):
            self.last_end = self.tokens[pos + 3].value
            l += 2
        elif self.have( pos + l, Token.PLUS ) or self.have( pos + l, Token.CONTINUOUS ):
            self.last_continuous = True
            l += 1

        if awards_are_events:
            if self.have( pos + l, 'Award', 'Award' ):
                self.last_name = self.tokens[pos + l].label
                l += 1
        else:
            if self.have( pos + l, 'Qualifier', 'Award' ):
                self.last_name = self.tokens[pos + l + 1].label + u' ' + self.tokens[pos + l].label
                l += 2

            if self.have( pos + l, 'Award' ):
                self.last_name = self.tokens[pos + l].label
                l += 1

            if self.have( pos + l, 'Qualifier' ):
                self.last_name = self.last_name + u' ' + self.tokens[pos + l].label
                l += 1

        if self.have( pos + l, Token.AT, 'Room' ):
            self.last_room = self.tokens[pos + l + 1].label
            l += 2

        del self.tokens[pos:pos + l]

        LOGGER.debug( 'Matched %s @ %s-%s %s in %s', self.last_name, self.last_start, self.last_end, self.last_continuous, self.last_room )
        return l

    def match_multiple_event_times( self, pos=0 ):
        self.match_initialize()

        if not self.have( pos, 'Event', 'Time', 'Time' ):
            return False

        self.last_name = self.tokens[pos].label
        self.last_actual = self.tokens[pos].label

        l = 2
        while self.have( pos + l, 'Time' ):
            l += 1

        self.time_list = [t.value for t in self.tokens[pos + 1: pos + l] ]
        del self.tokens[pos:pos + l]

        LOGGER.debug( 'Matched %s at %s', self.last_name, [ str( x ) for x in self.time_list ] )
        return l

    def match_room( self, pos=0 ):
        self.match_initialize()

        if not  self.have( pos, 'Room' ):
            return 0

        self.last_room = self.tokens[pos].label
        del self.tokens[pos]

        LOGGER.debug( 'Matched %s', self.last_room )
        return 1

    def match_room_events( self, pos=0 ):
        self.match_initialize()

        l = 0
        while self.have( pos + l, 'Event' ):
            l += 1

        if l and self.have( pos + l, 'Room' ):
            self.event_list = [t.label for t in self.tokens[pos:pos + l] ]
            self.last_room = self.tokens[pos + l].label

            l += 1
            del self.tokens[pos:pos + l]

            LOGGER.debug( 'Matched %s in %s', self.event_list, self.last_room )
            return l

        return 0

    def match_room_shift( self, pos=0 ):
        self.match_initialize()

        if not self.have( pos, 'Room', Token.SHIFT, 'Room' ):
            return 0

        l = 3
        self.default_room = self.tokens[pos].label
        self.shift_room = self.tokens[pos + 2].label

        if self.have( pos + l, Token.AT, 'Day', 'Time' ):
            self.shift_day = self.tokens[pos + l + 1]
            self.shift_time = self.tokens[pos + l + 2]
            l += 3

        del self.tokens[pos:pos + l]

        LOGGER.debug( 'Matched %s > %s @ %s:%s', self.default_room, self.shift_room, self.shift_day, self.shift_time )
        return l

#----- WBC Preview Schedule -------------------------------------------------

class WbcPreview( object ):
    """This class is used to parse schedule data from the annual preview pages"""

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
        'SLS': "Can't match 20 minute rounds",
        'LST': "Can't match 30 minute rounds, Can't handle 'until conclusion'",
        'KOT': "Can't match 30 minute rounds, Can't handle 'to conclusion'",
        'PGF': "Can't handle 'to conclusion'",

        'ADV': "Preview shows 1 hour for SF and F, not 2 hours per spreadsheet and pocket schedule",
        'BAR': "Pocket schedule shows R1@8/8:9, R2@8/8:14, R3@8/8:19, SF@8/9:9, F@8/9:14",
        'KFE': "Pocket schedule shows R1@8/6:9, R2@8/6:16, SF@8/7:9, F@8/7:16",
        'MED': "Preview shows heats taking place after SF/F",
        'RFG': "Should only have 1 demo, can't parse H1 room",
        'STA': "Conestoga is misspelled as Coonestoga",
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

    class Tourney( object ):
        """Class to organize events for a Preview tournament."""

        default_room = None
        shift_room = None
        shift_time = None
        draft_room = None

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

            self.tokenize_rooms( rows )
            self.tokenize_times( rows )

            if DEBUGGING:
                LOGGER.info( "%3s: rooms %s", self.code, Token.dump_list( self.room_tokens ) )
                LOGGER.info( "     times %s", Token.dump_list( self.event_tokens ) )

            self.parse_rooms()
            self.parse_events()

        def tokenize_times( self, rows ):
            """Parse 3rd row through next-to-last for event data"""

            self.event_tokens = []
            for row in rows[2:-1]:
                for td in row.findAll( 'td' ):
                    self.event_tokens.append( Token.START )
                    self.event_tokens += Token.tokenize( td )

        def tokenize_rooms( self, rows ):
            """"parse last row for room data"""

            self.room_tokens = Token.tokenize( rows[-1] )

        def parse_rooms( self ):

            # <default room>? [ SHIFT <shift room> [ AT <day> <time> ] ] { <event>+ <room> }*
            #
            # If this is PDT, then the shift room is special

            LOGGER.debug( 'Parsing rooms ...' )

            self.event_map = {}

            p = Parser( self.room_tokens )

            if p.match_room_shift():
                self.default_room = p.default_room
                if self.code == 'PDT':
                    self.draft_room = p.shift_room
                else:
                    self.shift_room = p.shift_room
                    self.shift_time = self.first_day + timedelta( days=p.shift_day.value ) + p.shift_time.value

            elif p.match_room():
                self.default_room = p.last_room

            while p.match_room_events():
                for e in p.event_list:
                    self.event_map[ e ] = p.last_room

            if p.count:
                LOGGER.warn( '%3s: Unmatched room tokens: %s', self.code, p.tokens )
                pass

        def parse_events( self ):

            # START <day> ( <event> <time> { PLUS | MINUS <end time> }? ( AT <room> )? )*
            # START <day> ( <day>* <event> <time> PLUS? )? <day>* <event> <time> PLUS?
            # [ AT <room> ] }+
            # If there are multiple days, then all of the events happen on all of the days (except demos)
            # If there are multiple times, then each represents a separate event on that day

            LOGGER.debug( 'Parsing events ...' )

            awards_are_events = self.code in ['UPF', 'VIP', 'WWR' ]
            two_weekends = self.code in [ 'BWD' ]

            self.events = []

            weekend_offset = 0
            weekend_start = 5 if two_weekends else 1

            p = Parser( self.event_tokens )

            if self.code in [ 'ATS']:
                pass

            while p.have_start():
                p.match( Token.START )

                days = []
                partial = []

                while p.count and not p.have_start():
                    if p.match_day():
                        # add day to day queue
                        day = p.last_day
                        days.append( day )

                        weekend_offset = 7 if day.value > weekend_start else weekend_offset

                        day_of_week = day.value + weekend_offset if day.value < 2 else day.value
                        midnight = self.first_day + timedelta( days=day_of_week )

                    elif p.match_multiple_event_times():
                        for etime in p.time_list:
                            # handle events immediately
                            dtime = midnight + etime
                            room = self.find_room( p.last_actual, dtime, p.last_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, TZ.localize( dtime ), room )
                            self.events.append ( e )

                    elif p.match_single_event_time( awards_are_events=awards_are_events ):
                        if p.last_name == 'Demo':
                            # handle demos immediately
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, p.last_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, TZ.localize( dtime ), room )
                            self.events.append ( e )

                        elif self.code == 'PDT' and p.last_name.endswith( 'FC' ):
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, self.draft_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name + u' Draft', TZ.localize( dtime ), room )
                            self.events.append ( e )

                            dtime = dtime + timedelta( hours=1 )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, TZ.localize( dtime ), self.default_room )
                            self.events.append ( e )

                        else:
                            # add event to event queue
                            e = WbcPreview.Event( self.code, self.name, p.last_name, p.last_start, p.last_room )
                            e.actual = p.last_actual
                            partial.append( e )
                    else:
                        p.recover()

                for day in days:
                    day_of_week = day.value + weekend_offset if day.value < 2 else day.value
                    midnight = self.first_day + timedelta( days=day_of_week )

                    for pevent in partial:
                        # add event to actual event list
                        dtime = midnight + pevent.time
                        room = self.find_room( pevent.actual, dtime, pevent.location )
                        e = WbcPreview.Event( self.code, self.name, pevent.type, TZ.localize( dtime ), room )
                        self.events.append( e )

            self.events.sort()

        def find_room( self, etype, etime, eroom ):
            room = self.shift_room if self.shift_room and etime >= self.shift_time else self.default_room
            room = self.event_map[ etype ] if self.event_map.has_key( etype ) else room
            room = eroom if eroom else room
            room = '-none-' if room == None else room

            if not room:
                pass

            return room

    def __init__( self, metadata, options, event_names ):
        self.meta = metadata
        self.options = options

        self.names = event_names  # mapping of codes to event names
        self.codes = event_names.keys()
        self.codes.sort()

        self.yy = self.options.year % 100
        if self.options.year != self.meta.this_year:
            self.PAGE_URL = "http://boardgamers.org/yearbkex%d/%%spge.htm" % ( self.yy, )

        LOGGER.info( 'Loading Preview schedule' )
        index = parse_url( self.INDEX_URL % ( self.yy, ) )
        if not index:
            LOGGER.error( 'Unable to load Preview index' )

        for option in index.findAll( 'option' ):
            value = option['value']
            if value == 'none' or value == '' or value == 'jnrpge.htm':
                continue
            pagecode = value[0:3]
            self.load_preview_page( pagecode )

        self.valid = True

    def load_preview_page( self, pagecode ):
        """Load and parse the preview page for a single tournament
        
        The schedule table within the page is a table that has a variable number of rows:
        
            [0] Contains the date the page was last updated -- ignored.
            [1] Contains the token code and other image codes
            [2:-2] Contains the schedule data, mostly as images, in two columns
            [-1] Contains the location information.
            
        As is the case with all of the WBC web pages, the HTML is ugly and malformed.
        """

        LOGGER.debug( 'Loading preview for %s', pagecode )

        # Map page codes to event codes
        code = self.codemap[ pagecode ] if self.codemap.has_key( pagecode ) else pagecode.upper()

#         # Skip any codes whose pages we can't handle
#         if self.skip.has_key( code ):
#             LOGGER.warn( 'Skipping %s: %s -- %s', code, self.names[ code ], self.skip[ code ] )
#             return

        if not self.names.has_key( code ):
#            LOGGER.error( "No event name for code [%s]; not loading preview", code )
            return

        # Load page
        url = self.PAGE_URL % pagecode
        page = parse_url( url )
        if not page:
            LOGGER.error( "Unable to load %s for [%s:%s]", url, pagecode, code )
            return

        t = WbcPreview.Tourney( code, self.names[ code ], page, self.meta.first_day )
        self.events[ code ] = t.events

#----- Schedule Comparison ---------------------------------------------------

class ScheduleComparer( object ):
    """This class knows enough about the different schedule sources to compare events"""

    TEMPLATE = 'report-template.html'

    def __init__( self, metadata, options, s, a, p=None ):
        self.options = options
        self.meta = metadata
        self.schedule = s
        self.allinone = a
        self.preview = p
        self.parser = None

    def verify_event_calendars( self ):
        """Compare the collections of events from both the calendars and the schedule"""

        LOGGER.info( 'Verifying event calendars against other sources' )

        schedule_key_set = set( self.schedule.current_tourneys )

        if self.allinone.valid:
            allinone_key_set = set( self.allinone.events.keys() )
            allinone_extras = allinone_key_set - schedule_key_set
            allinone_omited = schedule_key_set - allinone_key_set
        else:
            allinone_extras = set()
            allinone_omited = set()

        if self.preview.valid:
            preview_key_set = set( self.preview.events.keys() )
            preview_extras = preview_key_set - schedule_key_set
            preview_omited = schedule_key_set - preview_key_set
        else:
            preview_extras = set()
            preview_omited = set()

        add_space = False
        if len( allinone_extras ):
            LOGGER.error( 'Extra events present in All-in-One: %s', allinone_extras )
            add_space = True
        if len( allinone_omited ):
            LOGGER.error( 'Events omitted in All-in-One: %s', allinone_omited )
            add_space = True
        if len( preview_extras ):
            LOGGER.error( 'Extra events present in Preview: %s', preview_extras )
            add_space = True
        if len( preview_omited ):
            LOGGER.error( 'Events omitted in Preview: %s', preview_omited )
            add_space = True
        if add_space:
            LOGGER.error( '' )

        code_set = schedule_key_set
        if self.allinone.valid:
            code_set = code_set & allinone_key_set
        if self.preview.valid:
            code_set = code_set & preview_key_set

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
        prv_events = self.preview.events[ code ] if self.preview.valid else []
        prv_events = [ e for e in prv_events if e.type != 'Junior' ]
        cal_events = self.schedule.calendars[code].subcomponents

        # Find all of the unique times for any events
        ai1_timemap = dict( [ ( e.time.astimezone( TZ ), e ) for e in ai1_events ] )
        prv_timemap = dict( [ ( e.time.astimezone( TZ ), e ) for e in prv_events ] )
        cal_timemap = dict( [ ( e['dtstart'].dt.astimezone( TZ ), e ) for e in cal_events ] )
        time_set = set( ai1_timemap.keys() ) | set( prv_timemap.keys() ) | set( cal_timemap.keys() )
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

            # Fill in the Preview event, if present
            if prv_timemap.has_key( starting_time ):
                e = prv_timemap[ starting_time ]
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
        if self.preview.valid and self.preview.notes.has_key( code ) :
            discrepancies = True
            tr = self.parser.new_tag( 'tr' )
            td = self.parser.new_tag( 'td' )
            td['colspan'] = 8
            td['class'] = 'note'
            td.insert( 0, self.parser.new_string( self.preview.notes[ code ] ) )
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

        if self.preview.valid:
            a = self.parser.new_tag( 'a' )
            a['href'] = self.preview.PAGE_URL % code.lower()
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

        if not self.preview.valid:
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
                differences = set( [0, 1] )
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
        with codecs.open( path, 'w', 'utf-8' ) as f:
            f.write( self.parser.prettify() )

    @staticmethod
    def ai1_date_loc( ev ):
        """Generate a summary of an all-in-one event for comparer purposes"""

        start_time = ev.time.astimezone( TZ )
        location = ev.location
        location = 'Terrace' if location == 'Pt' else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M' ), location )

    @staticmethod
    def prv_date_loc( ev ):
        """Generate a summary of an preview event for comparer purposes"""

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

    # Parse the WBC Preview
    # names = dict( [( key_code, val_name ) for key_code, val_name in meta.names.items() if key_code in wbc_schedule.current_tourneys ] )
    names = dict( [( key_code, val_name ) for key_code, val_name in meta.names.items() ] )
    # names = { 'UPF': 'Up Front', 'PDT' : 'Pay Dirt', 'RDG' : 'Ra: The Dice Game' }
    wbc_preview = WbcPreview( meta, opts, names )

    # Compare the event calendars with the WBC All-in-One schedule and the preview
    comparer = ScheduleComparer( meta, opts, wbc_schedule, wbc_allinone, wbc_preview )
    comparer.verify_event_calendars()

    LOGGER.warn( "Done." )
