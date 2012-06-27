#! /usr/bin/env python

#----- Copyright (c) 2010 by W. Craig Trader ---------------------------------
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
# sudo pip install BeautifulSoup
# sudo pip install icalendar
# sudo pip install xlrd

#xxlint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0612,W0621,W0702,W0703
#pylint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0702

from bs4 import BeautifulSoup, Tag, NavigableString
from cgi import escape
from datetime import date, datetime, time, timedelta
from icalendar import Calendar, Event
from itertools import izip_longest
from optparse import OptionParser
import csv
import logging
import os
import pytz
import re
import unicodedata
import urllib2
import xlrd

logging.basicConfig( level=logging.WARN )
LOGGER = logging.getLogger( 'WBC' )
DEBUGGING = True

#----- Time Constants --------------------------------------------------------

TZ = pytz.timezone( 'America/New_York' ) # Tournament timezone
UTC = pytz.timezone( 'UTC' )             # UTC timezone (for iCal)

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

def process_options():
    """
    Parse command line options
    """

    LOGGER.debug( 'Parsing commandline options' )

    this_year = datetime.now( TZ ).year

    parser = OptionParser()
    parser.add_option( "-y", "--year", dest="year", metavar="YEAR", default=this_year, help="Year to process" )
    parser.add_option( "-t", "--type", dest="type", metavar="TYPE", default="xls", help="Type of file to process (csv,xls)" )
    parser.add_option( "-i", "--input", dest="input", metavar="FILE", default=None, help="Schedule spreadsheet to process" )
    parser.add_option( "-o", "--output", dest="output", metavar="DIR", default="build", help="Directory for results" )
    parser.add_option( "-v", "--verbose", dest="verbose", action="store_true", default=False )
    options, unused_args = parser.parse_args()

    return options

#----- WBC Event -------------------------------------------------------------

class WbcEvent( object ):
    """
    A WbcEvent encapsulates information about a single schedule line from the 
    WBC schedule spreadsheet.
    """

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

        self.type = ''
        self.rounds = 0
        self.freeformat = False
        self.grognard = False
        self.junior = False

        # read the data row using the subclass
        self.readrow( *args )

        # This test is for debugging purposes; string search on the spreadsheet event name
        if DEBUGGING and self.name.find( 'Ingenius' ) >= 0:
            pass

        # parse the data to generate useful fields
        self.checkrounds()
        self.checktypes( self.schedule.TYPES )
        self.checkrounds()
        self.checktypes( self.schedule.JUNIOR )
        self.checktimes()
        self.checkduration()
        self.checkcodes()

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
        return repr( self.__dict__ )

    def checkrounds( self ):
        """Check the current state of the event name to see if it describes a Heat or Round number"""

        match = re.search( r'([DHR]?)(\d+)/(\d+)$', self.name )
        if match:
            ( t, n, m ) = match.groups()
            text = match.group( 0 )
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

        if self.duration.endswith( "q" ):
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

class WbcCsvEvent( WbcEvent ):
    """This subclass of WbcEvent is used to parse CSV-formatted schedule data"""

    def __init__( self, *args ):
        WbcEvent.__init__( self, *args )

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

class WbcXlsEvent( WbcEvent ):
    """This subclass of WbcEvent is used to parse XLS-formatted schedule data"""

    def __init__( self, *args ):
        WbcEvent.__init__( self, *args )

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
                    val = datetime( *val ) # pylint: disable=W0142
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
                        raise ValueError( 'Unable to format ( % s ) as a time' % self.time )

        self.datetime = d.replace( hour=t.hour, minute=t.minute )

#----- WBC Meta Data ---------------------------------------------------------

class WbcMetadata( object ):
    """Load metadata about events that is not available from other sources"""

    # Data file names
    EVENTCODES = "wbc-event-codes.csv"
    OTHERCODES = "wbc-other-codes.csv"

    others = []         # List of non-tournament event matching data
    special = []        # List of non-tournament event codes
    tourneys = []       # List of tournament codes

    codes = {}          # Name -> Code map for events
    names = {}          # Code -> Name map for events

    durations = {}      # Special durations for events that have them
    grognards = {}      # Special durations for grognard events that have them
    playlate = {}       # Flag for events that may run past midnight

    first_day = None    # First calendar day for this year's convention

    def __init__( self ):
        self.load_tourney_codes()
        self.load_other_codes()

    def load_tourney_codes( self ):
        """
        Load all of the tourney codes (and alternate names) from their data file.
        """

        LOGGER.debug( 'Loading tourney event codes' )

        codefile = csv.DictReader( open( self.EVENTCODES ), delimiter=';' )
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

        codefile = csv.DictReader( open( self.OTHERCODES ), delimiter=';' )
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

    # Data file names
    TEMPLATE = "wbc-template.html"

    # Recognized event flags
    FLAVOR = [ 'FF', 'Circus', 'DDerby', 'Draft', 'Playoffs' ]
    JUNIOR = [ 'Jr', 'Jr.', 'Junior' ]
    TEEN = [ 'Teen' ]
    MENTOR = [ 'Mentoring' ]
    MULTIPLE = ['QF/SF/F', 'QF/SF', 'SF/F' ]
    SINGLE = [ 'QF', 'SF', 'F' ]
    STYLE = [ 'After Action Debriefing', 'After Action', 'Awards', 'Demo', 'Mulligan' ] + MULTIPLE + SINGLE

    TYPES = [ 'PC' ] + FLAVOR + JUNIOR + TEEN + MENTOR + STYLE

    rounds = {}         # Number of rounds for events that have rounds
    events = {}         # Events, grouped by code and then sorted by start date/time
    unmatched = []      # List of spreadsheet rows that don't match any coded events
    calendars = {}      # Calendars for each event code
    locations = {}      # Calendars by location
    dailies = {}        # Calendars by date

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
                event = WbcCsvEvent( self, i, row )
                self.categorize_event( event )
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
                print 'Reading row %d ...' % ( i + 1 )
            event = WbcXlsEvent( self, i + 1, header, sheet.row( i ), book.datemode )
            self.categorize_event( event )

    def categorize_event( self, event ):
        """
        Assign a spreadsheet entry to a matching event code.
        """
        if event.code:
            if self.events.has_key( event.code ):
                self.events[event.code].append( event )
            else:
                self.events[event.code] = [ event ]
        else:
            self.unmatched.append( event )

        self.meta.check_date( event.date )

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
        if DEBUGGING and event.code in ( 'SSB', ):
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
        9am the next morning.  There are two types of exceptions to this
        rule:  
        
        (1) some events run their finals directly after the semi-finals conclude,
            which can cause the final to run past midnight.
        (2) some events are scheduled to run multiple rounds past midnight.
        """

        # Calculate midnight, relative to the last event's start time
        midnight = datetime.fromordinal( start.toordinal() + 1 )

        # Calculate 9am tomorrow
        tomorrow = datetime( midnight.year, midnight.month, midnight.day, 9, 0, 0 )

        # Nominal start time and end time for the next event
        next_start = start + entry.length
        next_end = next_start + entry.length

        # Lookup the override code for this event
        playlate = self.meta.playlate.get( entry.code, None )

        if playlate == 'all':
            pass # always
        elif next_start > midnight:
            next_start = tomorrow
        elif next_end <= midnight:
            pass
        elif playlate == 'once':
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

        with open( os.path.join( self.options.output, "index.html" ), "w" ) as f:
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
            a = parser.new_tag( 'a' )
            a['class'] = 'eventlink'
            a['href'] = cls.safe_ics_filename( key )
            a.insert( 0, parser.new_string( escape( "%s" % label ) ) )
            td.insert( 0, span )
            td.insert( 1, a )
        else:
            td.insert( 0, parser.new_string( '&nbsp;' ) )
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
        li.insert( 0, parser.new_tag( 'a' ) )
        li.a['href'] = cls.safe_ics_filename( key )
        li.a.insert( 0, parser.new_string( escape( "%s" % label ) ) )
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
        same |= str( e1['summary'] ) == str( e2['summary'] )
        if altname:
            same |= str( e1['summary'] ) == altname
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
        &nbsp
        
    This is, frankly, horrible HTML.  But at least it's consistent, year-to-year, 
    and BeautifulSoup can parse it. 
    """

    SITE_URL = 'http://boardgamers.org/wbc/allin1.htm'

    TERRACE = 'Pt'

    colormap = {
        'green': 'Demo',
        '#07BED2': 'Mulligan',
        'red': 'Round',
        'blue': 'SF',
        '#AAAA00': 'F',
    }

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

    def __init__( self, metadata ):
        self.meta = metadata
        self.page = None

        self.load_table()

    def load_table( self ):
        """Parse the All-in-One schedule (HTML)"""

        LOGGER.info( 'Parsing WBC All-in-One schedule' )

        self.page = parse_url( self.SITE_URL )
        if not self.page:
            return

        tables = self.page.findAll( 'table' )
        rows = tables[1].findAll( 'tr' )
        for row in rows[1:]:
            self.load_row( row )

    def load_row( self, row ):
        """Parse an individual all-in-one row to find times and rooms for an event"""

        events = []

        cells = row.findAll( 'td' )
        code = str( cells[0].font.text ).strip( ';' )
        name = str( cells[1].font.text ).strip( ';* ' )

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
                        if day >= 32: # This only works because WBC is always in July/August, each of which has 31 days.
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
                                times, unused, entry = chunk.partition( ':' )
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

    SITE_URL = "http://boardgamers.org/yearbkex/%spge.htm"

    events = {}

    class Event( object ):
        """Simple data object to collect information about an event occuring at a specific time."""

        code = None
        name = None
        type = None
        time = None
        location = None

        def __cmp__( self, other ):
            return cmp( self.time, other.time )

        def __str__( self ):
            return '%s %s %s in %s at %s' % ( self.code, self.name, self.type, self.location, self.time )

    class Tourney( object ):
        """Simple data object to hold information about a group of events"""



    def __init__( self, metadata, event_names ):
        self.meta = metadata

        self.names = event_names # mapping of codes to event names
        self.codes = event_names.keys()
        self.codes.sort()

        LOGGER.info( 'Loading Yearbook schedule' )
        for code in self.codes:
            self.load_yearbook_page( code )

    def load_yearbook_page( self, code ):
        """Load and parse the yearbook page for a single tournament
        
        The schedule table within the page is a table that has a variable number of rows:
        
            [0] Contains the date the page was last updated -- ignored.
            [1] Contains the event code and other image codes
            [2:-2] Contains the schedule data, mostly as images, in two columns
            [-1] Contains the location information.
            
        As is the case with all of the WBC web pages, the HTML is ugly and malformed.
        """

#        LOGGER.warn( 'Loading yearbook for %s: %s', code, self.names[ code ] )

        # Load page
        pagecode = 'LVH' if code == 'LHV' else code
        page = parse_url( self.SITE_URL % pagecode.lower() )
        if not page:
            return

        # Find schedule / rows
        tables = page.findAll( 'table' )
        schedule = tables[2]
        rows = schedule.findAll( 'tr' )
#        LOGGER.warn( '%s: %d rows in schedule', code, len( rows ) )

        # parse last row for location data
        loc_data = rows[-1].find( 'center' )
        message = []
        for tag in loc_data.descendants:
            if isinstance( tag, NavigableString ):
                room = self.parse_room( tag )
                if room:
                    message.append( room )
            elif isinstance( tag, Tag ) and tag.name in ( 'img' ):
                event_type = self.parse_type( tag )
                message.append( event_type )
        print code, message

    def parse_room( self, ns ):
        """Remove extraneous characters and patterns from page text"""

        text = ns.strip()
        text = text.replace( '\n', '' )
        text = text.replace( u'\xa0', ' ' )
        text = text.replace( '#', 'Table ' )
        text = text.replace( ';', ',' )
        text = text.replace( ' ' * 8, ' ' )
        text = text.replace( ' ' * 4, ' ' )
        text = text.replace( ' ' * 2, ' ' )
        text = text.replace( ', Table', '' )
        return text

    def parse_type( self, tag ):
        event_type = tag['src'].lower()
        event_type = event_type.split( '/' )[-1]
        event_type = event_type.split( '.' )[0]
        event_type = 'demo' if event_type == 'demoweb' else event_type
        return event_type

#----- Schedule Comparison ---------------------------------------------------

class ScheduleComparer( object ):
    """This class knows enough about the different schedule sources to compare events"""

    def __init__( self, s, a, y=None ):
        self.schedule = s
        self.allinone = a
        self.yearbook = y

    def verify_event_calendars( self ):
        """Compare the collections of events from both the calendars and the schedule"""

        LOGGER.info( 'Verifying event calendars against All-in-One schedule' )

        schedule_key_set = set( self.schedule.current_tourneys )
        allinone_key_set = set( self.allinone.events.keys() )
        if self.yearbook:
            yearbook_key_set = set( self.yearbook.events.keys() )

        allinone_extras = allinone_key_set - schedule_key_set
        allinone_omited = schedule_key_set - allinone_key_set
        if self.yearbook:
            yearbook_omited = schedule_key_set - yearbook_key_set
        else:
            yearbook_omited = set()

        if len( allinone_extras ):
            LOGGER.error( 'Extra events present in All-in-One: %s', allinone_extras )
        if len( allinone_omited ):
            LOGGER.error( 'Events omitted in All-in-One: %s', allinone_omited )
        if len( yearbook_omited ):
            LOGGER.error( 'Events omitted in Yearbook: %s', yearbook_omited )
        if len( allinone_extras ) + len( allinone_omited ) + len( yearbook_omited ):
            LOGGER.error( '' )

        code_set = schedule_key_set & allinone_key_set
        if self.yearbook:
            code_set = code_set & yearbook_key_set
        codes = list( code_set )
        codes.sort()

        for code in codes:
            self.check_schedule_against_allinone( code )

    def check_schedule_against_allinone( self, code ):
        """Compare the generated calendar for an event against the all-in-one schedule 
        for that event.  If there are differences, log them."""

        ai1_events = self.allinone.events[ code ]
        cal_events = self.schedule.calendars[code].subcomponents

        # Remove calendar events that aren't tracked on the all-in-one schedule
        cal_events = [ e for e in cal_events if not e['summary'].find( ' Jr' ) >= 0 ]
#       cal_events = [ e for e in cal_events if not e['summary'].endswith( 'After Action' ) ]
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
            LOGGER.error( '%s: All-in-One________________________ Calendar__________________________ __Length Name________________', code )
            cal_other = [ '%8s %s' % ( x['duration'].dt, x['summary'] ) for x in cal_events ]
            for tab, cal, other in izip_longest( ai1_comparison, cal_comparison, cal_other, fillvalue='' ):
                mark = ' ' if tab == cal else '*'
                LOGGER.error( '%4s %-34s %-34s %s', mark, tab, cal, other )
            LOGGER.error( '' )

    @staticmethod
    def ai1_date_loc( ev ):
        """Generate a summary of an all-in-one event for comparer purposes"""

        start_time = ev.time.astimezone( TZ )
        location = ev.location
        location = 'Terrace' if location == 'Pt' else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M:%S' ), location )

    @staticmethod
    def cal_date_loc( sc ):
        """Generate a summary of a calendar event for comparer purposes"""

        start_time = sc['dtstart'].dt.astimezone( TZ )
        location = sc['location']
        location = 'Terrace' if location.startswith( 'Terrace' ) else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M:%S' ), location )

#----- Real work happens here ------------------------------------------------

if __name__ == '__main__':

    meta = WbcMetadata()

    opts = process_options()

    # Load a schedule from a spreadsheet, based upon commandline options.
    wbc_schedule = WbcSchedule( meta, opts )

#    testdata = { '8XX': '18XX', 'AFK': 'Afrika Korps', 'TTN': 'Titan' }
#    wbc_yearbook = WbcYearbook( meta, testdata )

#if False:

    # Create calendar events from all of the spreadsheet events.
    wbc_schedule.create_wbc_calendars()

    # Write the individual event calendars.
    wbc_schedule.write_all_calendar_files()

    # Build the HTML index.
    wbc_schedule.write_index_page()

    # Print the unmatched events for rework.
    wbc_schedule.report_unprocessed_events()

    # Parse the WBC All-in-One schedule
    wbc_allinone = WbcAllInOne( meta )

    # Parse the WBC Yearbook
    # names = dict( [( kx, vx ) for kx, vx in meta.names.items() if kx in wbc_schedule.current_tourneys ] )
    # wbc_yearbook = WbcYearbook( meta, names )

    # Compare the event calendars with the WBC All-in-One schedule and the yearbook
    comparer = ScheduleComparer( wbc_schedule, wbc_allinone )
    comparer.verify_event_calendars()

    LOGGER.info( "Done." )
