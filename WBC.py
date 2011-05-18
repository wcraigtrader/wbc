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

# These are the 'non-standard' libraries we need ...
#
# sudo apt-get install pip
# sudo pip install pytz
# sudo pip install BeautifulSoup
# sudo pip install icalendar
# sudo pip install xlrd

from BeautifulSoup import BeautifulSoup, Tag, NavigableString
from datetime import date, datetime, time, timedelta
from icalendar import Calendar, Event
from optparse import OptionParser
from pytz import timezone
import csv
import os
import re
import sys
import unicodedata
import xlrd

#----- WBC Event -------------------------------------------------------------

class WbcEvent( object ):

    def __init__( self, schedule, line, *args ):

        self.schedule = schedule
        self.line = line

        # default values for calculated fields
        self.type = ''
        self.rounds = 0
        self.freeformat = False

        # read the data row using the subclass
        self.readrow( *args )

        if self.name.find( 'Puerto Rico' ) >= 0:
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

    def checkcodes( self ):
        self.code = None

        # Check for tournament codes
        if self.schedule.codes.has_key( self.name ):
            self.code = self.schedule.codes[ self.name ]
            self.name = self.schedule.names[ self.code ]

            # If this event has rounds, save them for later use
            if self.rounds:
                self.schedule.rounds[ self.code ] = self.rounds

        else:
            # Check for non-tournament groupings
            for o in self.schedule.others:
                if ( ( o['format'] and o['format'] == self.format ) or
                    ( o['name'] and o['name'] == self.name ) ):
                    self.code = o['code']
                    return

    def checkduration( self ):
        if self.__dict__.has_key( 'continuous' ):
            self.continuous = ( self.continuous in ( 'C', 'Y' ) )
        else:
            self.continuous = False

        if self.duration.endswith( "q" ):
            self.continuous = True
            self.duration = self.duration[:-1]

        if self.duration == '<1':
           self.duration = 0.5

        if self.duration == '2/1':
           self.duration = 1

        if self.duration and self.duration != '-':
            self.length = timedelta( minutes=60 * float( self.duration ) )

    def checkrounds( self ):
        match = re.search( r'([HR]?)(\d+)/(\d+)$', self.name )
        if match:
            ( t, n, m ) = match.groups()
            text = match.group( 0 )
            if t == "R":
                self.start = int( n )
                self.rounds = int( m )
                self.name = self.name[:-len( text )].strip()
            elif t == "H" or t == '':
                self.type = self.type + ' ' + text
                self.type = self.type.strip()
                self.name = self.name[:-len( text )].strip()

    def checktypes( self, types ):
        for type in types:
            if self.name.endswith( type ):
                self.type = type + ' ' + self.type
                self.name = self.name[:-len( type )].strip()
                if type == 'FF':
                    self.freeformat = True

        self.type = self.type.strip()

#----- WBC Event (read from CSV spreadsheet) ---------------------------------

class WbcCsvEvent( WbcEvent ):
    def __init__( self, *args ):
        WbcEvent.__init__( self, *args )

    def readrow( self, *args ):
        row = args[0]

        for ( key, val ) in row.items():
            self.__setattr__( key, val )

        self.name = self.event.strip()

    def checktimes( self ):
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

    def __init__( self, *args ):
        WbcEvent.__init__( self, *args )

    def readrow( self, *args ):
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
                val = unicodedata.normalize( 'NFKD', val.value ).encode( 'ascii', 'ignore' )
            elif val.ctype == xlrd.XL_CELL_NUMBER:
                val = str( int( val.value ) )
            elif val.ctype == xlrd.XL_CELL_DATE:
                val = xlrd.xldate_as_tuple( val.value, datemode )
                if val[0]:
                    val = datetime( *val )
                else:
                    val = time( val[3], val[4], val[5] )
            else:
                raise ValueError( "Unhandled Excel cell type (%s) for %s" % ( val.ctype, key ) )

            self.__setattr__( key, val )

        self.name = self.event.strip()

    def checktimes( self ):
        d = self.date

        if self.time.__class__ is time:
            t = self.time
        else:
            try:
                t = int( self.time )
                if t > 23:
                    t = time( 23, 59 )
                    self.length = timedelta( minutes=1 )
                else:
                    t = time( t )
            except Exception as e1:
                try:
                    t = datetime.strptime( self.time, "%H:%M" )
                except Exception as e2:
                    try:
                        t = datetime.strptime( self.time, "%I:%M:%S %p" )
                    except Exception as e3:
                        raise ValueError( 'Unable to format (%s) as a time' % self.time )

        self.datetime = d.replace( hour=t.hour, minute=t.minute )

#----- WBC Schedule ----------------------------------------------------------

class WbcSchedule( object ):
    """
    The WbcSchedule class parses the entire WBC schedule and creates 
    iCalendar calendars for each event (with vEvents for each time slot).
    """

    timezone = timezone( 'US/Eastern' )    # Tournament timezone

    # Data file names
    EVENTCODES = "wbc-event-codes.csv"
    OTHERCODES = "wbc-other-codes.csv"
    TEMPLATE = "wbc-template.html"

    # Recognized event flags
    FLAVOR = [ 'FF', 'Circus', 'DDerby', 'Draft', 'Playoffs' ]
    JUNIOR = [ 'Jr', 'Jr.', 'Junior' ]
    TEEN = [ 'Teen' ]
    MENTOR = [ 'Mentoring' ]
    MULTIPLE = ['QF/SF/F', 'QF/SF', 'SF/F' ]
    SINGLE = [ 'QF', 'SF', 'F' ]
    STYLE = [ 'After Action', 'Demo', 'Mulligan' ] + MULTIPLE + SINGLE

    TYPES = [ 'PC' ] + FLAVOR + JUNIOR + TEEN + MENTOR + STYLE

    others = []         # List of non-tournament event matching data
    special = []        # List of non-tournament event codes
    tourneys = []       # List of tournament codes

    codes = {}          # Name -> Code map for events
    names = {}          # Code -> Name map for events

    rounds = {}         # Number of rounds for events that have rounds
    events = {}         # Events, grouped by code and then sorted by start date/time
    unmatched = []      # List of spreadsheet rows that don't match any coded events
    calendars = {}      # Calendars for each event code
    locations = {}      # Calendars by location
    dailies = {}        # Calendars by date

    def __init__( self ):
        """
        Initialize a schedule
        """
        self.processed = datetime.now( self.timezone )

        self.process_options()
        self.load_tourney_codes()
        self.load_other_codes()
        self.load_events()

        self.prodid = "WBC %s" % self.options.year

    def process_options( self ):
        """
        Parse command line options
        """

        parser = OptionParser()
        parser.add_option( "-y", "--year", dest="year", metavar="YEAR", default=self.processed.year, help="Year to process" )
        parser.add_option( "-t", "--type", dest="type", metavar="TYPE", default="xls", help="Type of file to process (csv,xls)" )
        parser.add_option( "-i", "--input", dest="input", metavar="FILE", default=None, help="Schedule spreadsheet to process" )
        parser.add_option( "-o", "--output", dest="output", metavar="DIR", default="test", help="Directory for results" )
        parser.add_option( "-v", "--verbose", dest="verbose", action="store_true", default=False )
        self.options, args = parser.parse_args()

        if not os.path.exists( self.options.output ):
            os.makedirs( self.options.output )

    def load_tourney_codes( self ):
        """
        Load all of the tourney codes (and alternate names) from their data file.
        """
        codefile = csv.DictReader( open( self.EVENTCODES ), delimiter=';' )
        for row in codefile:
            c = row['Code'].strip()
            n = row['Name'].strip()
            self.codes[ n ] = c
            self.names[ c ] = n
            self.tourneys.append( c )

            for altname in [ 'Alt1', 'Alt2', 'Alt3', 'Alt4', 'Alt5', 'Alt6']:
                if row[altname]:
                    a = row[altname].strip()
                    self.codes[a] = c

    def load_other_codes( self ):
        """
        Load all of the non-tourney codes from their data file.
        """
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

    def load_events( self ):
        """
        Process all of the events in the spreadsheet
        """
        print "Scanning schedule spreadsheet..."
        filename = self.options.input
        if self.options.type == 'csv':
            if not filename:
                filename = "wbc-%s-schedule.csv" % self.processed.year
            self.scan_csv_file( filename )
        elif self.options.type == 'xls':
            if not filename:
                filename = "WBCScheduleSpreadsheet%s.xls" % self.processed.year
            self.scan_xls_file( filename )

    def scan_csv_file( self, filename ):
        """
        Read a CSV-formatted file and generate WBC events for each row
        """
        eventfile = csv.DictReader( open( filename ) , delimiter=';' )
        i = 2
        for row in eventfile:
            event = WbcCsvEvent( self, i, row )
            self.categorize_event( event )
            i = i + 1

    def scan_xls_file( self, filename ):
        """
        Read an Excel spreadsheet and generate WBC events for each row.
        """
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

    def process_wbc_events( self ):
        """
        Process all of the spreadsheet entries, by event code, then by time.
        """
        for code, list in self.events.items():
            list.sort( lambda x, y: cmp( x.datetime, y.datetime ) )
#            if code == 'PRO':
#                pass
            for event in list:
                self.process_event( event )

    def process_event( self, event ):
        """
        For a spreadsheet entry, generate calendar events as follows:
        
        If the entry has rounds, add calendar events for each round.
        If the entry is marked as continuous, and the code is WAW, 
           it gets special handling.
        If the entry is marked as continuous, and it's marked 'HMSE', 
           there's no clue as to how many actual heats there are, so just add one event.
        If the entry is marked as continuous, add as many events as there are types coded.
        Otherwise, 
           add a single event for the entry.
        """

        calendar = self.get_or_create_event_calendar( event.code )

        if event.rounds:
            self.process_event_with_rounds( calendar, event )
        elif event.continuous and event.code == 'WAW':
            self.process_all_week_event( calendar, event )
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

    def process_continuous_event( self, calendar, event ):
        """
        Process multiple back-to-back events that are not rounds, per se.
        """
        start = event.datetime
        for type in event.type.split( '/' ):
            name = event.name + ' ' + type
            alternative = self.alternate_round_name( event, type )
            self.add_event( calendar, event, start=start, name=name, altname=alternative )
            start = start + event.length

    def process_event_with_rounds( self, calendar, event ):
        """
        Process multiple back-to-back rounds
        """
        start = event.datetime
        name = event.name + ' ' + event.type
        name = name.strip()

        rounds = range( int( event.start ), int( event.rounds ) + 1 )
        for r in rounds:
            duration = event.length
            if event.freeformat and event.format == 'SwEl':
                midnight = midnight = start.date() + timedelta( days=1 )
                duration = datetime( midnight.year, midnight.month, midnight.day ) - start

            label = "%s R%s/%s" % ( name, r, event.rounds )
            self.add_event( calendar, event, start=start, duration=duration, name=label )

            # Check for rounds that would begin after midnight
            next = start + duration
            if next.toordinal() > start.toordinal():
                start = datetime( next.year, next.month, next.day, 9, 0, 0 )
            else:
                start = next

    def process_all_week_event( self, calendar, event ):
        """
        Process an event that runs contiuously all week long.
        """
        start = event.datetime
        remaining = event.length
        while ( remaining.days or remaining.seconds ):
            midnight = midnight = start.date() + timedelta( days=1 )
            duration = datetime( midnight.year, midnight.month, midnight.day ) - start
            if duration > remaining:
                duration = remaining

            self.add_event( calendar, event, start=start, duration=duration, replace=False )

            start = datetime( midnight.year, midnight.month, midnight.day, 9, 0, 0 )
            remaining = remaining - duration

    def alternate_round_name( self, event, type=None ):
        """
        Create the equivalent round name for a given event.
        
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
        if the current event type is 'QF', 'SF', or 'F'.   
        """
        type = type if type else event.type

        alternative = None
        if self.rounds.has_key( event.code ) and type in self.SINGLE:
            r = self.rounds[ event.code ]
            offset = ( len( self.SINGLE ) - self.SINGLE.index( type ) ) - 1
            alternative = "%s R%s/%s" % ( event.name, r - offset, r )
        return alternative

    def get_or_create_event_calendar( self, code ):
        """
        For a given event code, return the iCalendar that matches that code.
        If there is no pre-existing calendar, create a new one.
        """
        if self.calendars.has_key( code ):
            return self.calendars[ code ]

        calendar = Calendar()
        calendar.add( 'VERSION', '2.0' )
        calendar.add( 'PRODID', '-//%s %s//ct7//' % ( self.prodid, code ) )
        calendar.add( 'SUMMARY', self.names[ code ] )
        calendar.add( 'DESCRIPTION', "%s %s: %s" % ( self.prodid, code, self.names[code] ) )

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

        calendar = Calendar()
        calendar.add( 'VERSION', '2.0' )
        calendar.add( 'PRODID', '-//%s %s//ct7//' % ( self.prodid, location ) )
        calendar.add( 'SUMMARY', 'Events in ' + location )
        calendar.add( 'DESCRIPTION', "%s: Events in %s" % ( self.prodid, location ) )

        self.locations[ location ] = calendar

        return calendar

    def get_or_create_daily_calendar( self, date ):
        """
        For a given date, return the iCalendar that matches that date.
        If there is no pre-existing calendar, create a new one.
        """
        key = date.dt.date()
        name = date.dt.strftime( '%A, %B %d' )

        if self.dailies.has_key( key ):
            return self.dailies[ key ]

        calendar = Calendar()
        calendar.add( 'VERSION', '2.0' )
        calendar.add( 'PRODID', '-//%s %s//ct7//' % ( self.prodid, key ) )
        calendar.add( 'SUMMARY', 'Events on ' + name )
        calendar.add( 'DESCRIPTION', '%s: Events on %s' % ( self.prodid, name ) )

        self.dailies[ key ] = calendar

        return calendar

    def add_event( self, calendar, entry, start=None, duration=None, name=None, altname=None, replace=True ):
        """
        Add a new vEvent to the given iCalendar for a given spreadsheet entry.
        """
        name = name if name else entry.name
        start = start if start else entry.datetime
        duration = duration if duration else entry.length

        e = Event()
        e.add( 'SUMMARY', name )
        e.add( 'DTSTART', self.timezone.localize( start ) )
        e.add( 'DURATION', duration )
        e.add( 'LOCATION', entry.location )
        e.add( 'ORGANIZER', entry.gm )

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
        print "Unprocessed entries ..."
        self.unmatched.sort( cmp=lambda x, y: cmp( x.name, y.name ) )
        for event in self.unmatched:
            print "Row %3d [%s] %s" % ( event.line, event.name, event )

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
        print "Saving calendars..."

        # Create bulk calendars
        everything = Calendar()
        everything.add( 'VERSION', '2.0' )
        everything.add( 'PRODID', '-//' + self.prodid + ' Everything//ct7//' )
        everything.add( 'SUMMARY', 'WBC %s All-in-One Schedule' % self.options.year )

        tournaments = Calendar()
        tournaments.add( 'VERSION', '2.0' )
        tournaments.add( 'PRODID', '-//' + self.prodid + ' Tournaments//ct7//' )
        tournaments.add( 'SUMMARY', 'WBC %s Tournaments Schedule' % self.options.year )

        # For all of the event calendars
        for code, calendar in self.calendars.items():

            # Write the calendar itself
            self.write_calendar_file( calendar, code )

            # Add all calendar events to the master calendar
            everything.subcomponents += calendar.subcomponents

            # Add all the tourney events to the tourney calendar
            if code in self.tourneys:
                tournaments.subcomponents += calendar.subcomponents

            # For each calendar event
            for event in calendar.subcomponents:

                # Add it to the appropriate location calendar 
                location = self.get_or_create_location_calendar( event['LOCATION'] )
                location.subcomponents.append( event )

                # Add it to the appropriate daily calendar
                daily = self.get_or_create_daily_calendar( event['DTSTART'] )
                daily.subcomponents.append( event )

        # Write the master and tourney calendars
        self.write_calendar_file( everything, "all-in-one" )
        self.write_calendar_file( tournaments, "tournaments" )

        # Write the location calendars
        for location, calendar in self.locations.items():
            self.write_calendar_file( calendar, location )

        # Write the daily calendars
        for date, calendar in self.dailies.items():
            self.write_calendar_file( calendar, date )

    def write_index_page( self ):
        """
        Using an HTML Template, create an index page that lists 
        all of the created calendars.
        """
        print "Saving index page..."
        with open( self.TEMPLATE, "r" ) as f:
            template = f.read()

        index = BeautifulSoup( template )

        # Locate insertion points
        title = index.find( 'title' )
        header = index.find( 'div', { 'id' : 'header' } )
        footer = index.find( 'div', { 'id' : 'footer' } )
        event_list = index.find( 'div', { 'id' : 'tournaments' } ).ol
        other_list = index.find( 'div', { 'id' : 'other' } ).ul
        every_list = index.find( 'div', { 'id' : 'special' } ).ul
        daily_list = index.find( 'div', { 'id' : 'daily' } ).ul
        place_list = index.find( 'div', { 'id' : 'location' } ).ul

        # Page title
        title.insert( 0, NavigableString( "WBC %s Event Schedule" % self.options.year ) )
        header.h1.insert( 0, NavigableString( "WBC %s Event Schedule" % self.options.year ) )
        footer.p.insert( 0, NavigableString( "Updated on %s" % self.processed.strftime( "%A, %d %B %Y %H:%M %Z" ) ) )

        # All-in-One calendar
        line = Tag( index, 'li' )
        line.insert( 0, Tag( index, 'a' ) )
        line.a['href'] = self.safe_ics_filename( "all-in-one" )
        line.a.insert( 0, NavigableString( 'WBC %s All-in-One Schedule' % self.options.year ) )
        every_list.insert( len( every_list ), line )

        # All tournaments calendar
        line = Tag( index, 'li' )
        line.insert( 0, Tag( index, 'a' ) )
        line.a['href'] = self.safe_ics_filename( "tournaments" )
        line.a.insert( 0, NavigableString( 'WBC %s Tournaments Schedule' % self.options.year ) )
        every_list.insert( len( every_list ), line )

        # Location calendars
        keys = self.locations.keys()
        keys.sort()
        for location in keys:
            calendar = self.locations[location]
            line = Tag( index, 'li' )
            line.insert( 0, Tag( index, 'a' ) )
            line.a['href'] = self.safe_ics_filename( location )
            line.a.insert( 0, NavigableString( '%s Schedule' % location ) )
            place_list.insert( len( place_list ), line )

        # Daily calendars
        keys = self.dailies.keys()
        keys.sort()
        for date in keys:
            calendar = self.dailies[date]
            line = Tag( index, 'li' )
            line.insert( 0, Tag( index, 'a' ) )
            line.a['href'] = self.safe_ics_filename( date )
            line.a.insert( 0, NavigableString( '%s Schedule' % date.strftime( '%A, %B %d' ) ) )
            daily_list.insert( len( daily_list ), line )

        # Individual calendars
        calendar_list = self.calendars.items()
        calendar_list.sort( cmp=lambda x, y: cmp( x[1]['SUMMARY'], y[1]['SUMMARY'] ) )
        for code, calendar in calendar_list:
            line = Tag( index, 'li' )
            if code in self.special:
                line.insert( 0, Tag( index, 'a' ) )
                line.a['href'] = self.safe_ics_filename( code )
                line.a.insert( 0, NavigableString( calendar['summary'] ) )
                other_list.insert( len( other_list ), line )
            else:
                line.insert( 0, Tag( index, 'span' ) )
                line.span['class'] = 'eventcode'
                line.span.insert( 0, NavigableString( self.safe_html( code ) + ': ' ) )
                line.insert( 1, Tag( index, 'a' ) )
                line.a['href'] = self.safe_ics_filename( code )
                line.a['class'] = 'eventlink'
                line.a.insert( 0, NavigableString( calendar['summary'] ) )
                event_list.insert( len( event_list ), line )

        print "Saving index page..."
        with open( os.path.join( self.options.output, "index.html" ), "w" ) as f:
            f.write( index.prettify() )

    @classmethod
    def serialize_calendar( cls, calendar ):
        """This fixes portability quirks in the iCalendar library:
        1) The vCalendar's 'VERSION:2.0' line should appear directly after the 'BEGIN:VCALENDAR' line;
           The iCalendar library sorts all its properties alphabetically, which violates this.
        2) The iCalendar library generates event start date/times as 'DTSTART;DATE=VALUE:yyyymmddThhmmssZ';
           the more acceptable format is 'DTSTART:yyyymmddThhmmssZ'
        3) The iCalendar library doesn't sort the events in a given calendar by date/time.
        """
        c = calendar
        c.subcomponents.sort( cmp=cls.compare_icalendar_events )

        s = c.as_string()
        s = s.replace( "\r\nVERSION:2.0", "" )
        s = s.replace( "BEGIN:VCALENDAR", "BEGIN:VCALENDAR\r\nVERSION:2.0" )
        s = s.replace( ";VALUE=DATE:", ":" )
        return s

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

    @staticmethod
    def safe_html( name ):
        """
        HTML escaping for Dummies.
        """
        name = name.replace( '&', '&amp;' )
        name = name.replace( '<', '&lt;' )
        name = name.replace( '>', '&gt;' )
        return name

#----- WBC YearBook ----------------------------------------------------------

class WbcYearBook( object ):
    pass

#----- Real work happens here ------------------------------------------------

if __name__ == '__main__':

    # Load a schedule from a spreadsheet, based upon commandline options.
    schedule = WbcSchedule()

    # Create calendar events from all of the spreadsheet events.
    schedule.process_wbc_events()

    # Write the individual event calendars.
    schedule.write_all_calendar_files()

    # Build the HTML index.
    schedule.write_index_page()

    # Print the unmatched events for rework.
    schedule.report_unprocessed_events()

    print "Done."
