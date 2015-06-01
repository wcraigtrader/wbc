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

from bs4 import BeautifulSoup
import codecs
import logging
import os

from WbcUtility import asLocal


LOGGER = logging.getLogger( 'WbcScheduleComparison' )

#----- Schedule Comparison ---------------------------------------------------

class ScheduleComparer( object ):
    """This class knows enough about the different schedule sources to compare events"""

    TEMPLATE = 'report-template.html'

    def __init__( self, metadata, s, a, p=None ):
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

        if len( allinone_extras ):
            LOGGER.error( 'Extra events present in All-in-One: %s', allinone_extras )
        if len( allinone_omited ):
            LOGGER.error( 'Events omitted in All-in-One: %s', allinone_omited )
        if len( preview_extras ):
            LOGGER.error( 'Extra events present in Preview: %s', preview_extras )
        if len( preview_omited ):
            LOGGER.error( 'Events omitted in Preview: %s', preview_omited )

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
        if self.meta.fullreport:
            text = "WBC %s Schedule Details" % self.meta.year
        else:
            text = "WBC %s Schedule Discrepancies" % self.meta.year

        title.insert( 0, self.parser.new_string( text ) )
        header.h1.insert( 0, self.parser.new_string( text ) )
        footer.p.insert( 0, self.parser.new_string( "Updated on %s" % self.meta.now.strftime( "%A, %d %B %Y %H:%M %Z" ) ) )

    def report_discrepancies( self, code ):
        """Format the discrepancies for a given tournament"""

        # Find all of the matching events from each schedule
        ai1_events = self.allinone.events[ code ] if self.allinone.valid else []
        prv_events = self.preview.events[ code ] if self.preview.valid else []
        prv_events = [ e for e in prv_events if e.type != 'Junior' ]
        cal_events = self.schedule.calendars[code].subcomponents

        # Find all of the unique times for any events
        ai1_timemap = dict( [ ( asLocal( e.time ), e ) for e in ai1_events ] )
        prv_timemap = dict( [ ( asLocal( e.time ), e ) for e in prv_events ] )
        cal_timemap = dict( [ ( asLocal( e.decoded( 'dtstart' ) ), e ) for e in cal_events ] )
        time_set = set( ai1_timemap.keys() ) | set( prv_timemap.keys() ) | set( cal_timemap.keys() )
        time_list = list( time_set )
        time_list.sort()

        label = self.meta.names[code]

        rows = []
        discrepancies = False

        self.create_discrepancy_header( rows, code )

        # For each date/time combination, compare all of the events at that time
        for starting_time in time_list:
            # Start with empty cells
            details = [( None, ), ( None, ), ( None, ), ]

            # Fill in the Preview event, if present
            if prv_timemap.has_key( starting_time ):
                e = prv_timemap[ starting_time ]
                location = 'Terrace' if e.location and e.location.startswith( 'Terr' ) else e.location
                details[0] = ( location, e.type )

            # Fill in the All-in-One event, if present
            if ai1_timemap.has_key( starting_time ):
                e = ai1_timemap[ starting_time ]
                location = 'Terrace' if e.location == 'Pt' else e.location
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
        if discrepancies or self.meta.fullreport:
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

        if self.preview.valid:
            a = self.parser.new_tag( 'a' )
            a['href'] = self.meta.url[ code ]
            a.insert( 0, self.parser.new_string( 'Event Preview' ) )
            th = self.parser.new_tag( 'th' )
            th['colspan'] = 2
            th.insert( 0, a )
            tr.insert( len( tr ), th )

        if self.allinone.valid:
            th = self.parser.new_tag( 'th' )
            th['colspan'] = 2
            th.insert( 0, self.parser.new_string( 'All-in-One' ) )
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

        path = os.path.join( self.meta.output, "report.html" )
        with codecs.open( path, 'w', 'utf-8' ) as f:
            f.write( self.parser.prettify() )

    def ai1_date_loc( self, ev ):
        """Generate a summary of an all-in-one event for comparer purposes"""

        start_time = asLocal( ev.time )
        location = ev.location
        location = 'Terrace' if location == 'Pt' else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M' ), location )

    def prv_date_loc( self, ev ):
        """Generate a summary of an preview event for comparer purposes"""

        start_time = asLocal( ev.time )
        location = ev.location
        location = 'Terrace' if location.startswith( 'Terr' ) else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M' ), location )

    def cal_date_loc( self, sc ):
        """Generate a summary of a calendar event for comparer purposes"""

        start_time = asLocal( sc.decoded( 'dtstart' ) )
        location = sc['location']
        location = 'Terrace' if location.startswith( 'Terrace' ) else location
        return '%s : %s' % ( start_time.strftime( '%a %m-%d %H:%M' ), location )

