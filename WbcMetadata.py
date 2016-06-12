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

from datetime import datetime
from optparse import OptionParser
import csv
import logging
import os
import re

from bs4 import Tag, NavigableString
from WbcUtility import parse_url, TZ, normalize

LOGGER = logging.getLogger( 'WbcMetaData' )


# ----- WBC Meta Data ---------------------------------------------------------

class WbcMetadata( object ):
    """Load metadata about events that is not available from other sources"""

    now = datetime.now( TZ )
    this_year = now.year

    # Data file names
    EVENTCODES = os.path.join( "meta", "wbc-event-codes.csv" )
    OTHERCODES = os.path.join( "meta", "wbc-other-codes.csv" )

    SITE_URL = "http://boardgamers.org/"
    PREVIEW_INDEX_URL = SITE_URL + "previews_%d.html"

    # Bad code in event preview index -> actual event code
    MISCODES = {}  # {'gmb': 'GBM',}

    others = []  # List of non-tournament event matching data
    special = []  # List of non-tournament event codes
    tourneys = []  # List of tournament codes

    codes = {}  # Name -> Code map for events
    names = {}  # Code -> Name map for events

    durations = {}  # Special durations for events that have them
    grognards = {}  # Special durations for grognard events that have them
    playlate = {}  # Flag for events that may run past midnight
    url = {}  # Code -> URL for event preview for this event code

    first_day = None  # First calendar day for this year's convention

    year = this_year  # Year to process
    type = "xls"  # Type of spreadsheet to parse
    input = None  # Name of spreadsheet to parse
    output = "site"  # Name of directory for results
    write_files = True  # Whether or not to output files
    fullreport = False  # True for more detailed discrepancy reporting
    verbose = False  # True for more detailed logging
    debug = False  # True for debugging and even more detailed logging

    def __init__(self):
        self.process_options( )
        self.load_tourney_codes( )
        self.load_other_codes( )
        self.load_preview_index( )

    def process_options(self):
        """
        Parse command line options
        """

        parser = OptionParser( )
        parser.add_option( "-y", "--year", dest="year", metavar="YEAR", default=self.this_year, help="Year to process" )
        parser.add_option( "-t", "--type", dest="type", metavar="TYPE", default="xls",
                           help="Type of file to process (csv,xls)" )
        parser.add_option( "-i", "--input", dest="input", metavar="FILE", default=None,
                           help="Schedule spreadsheet to process" )
        parser.add_option( "-o", "--output", dest="output", metavar="DIR", default="build",
                           help="Directory for results" )
        parser.add_option( "-f", "--full-report", dest="fullreport", action="store_true", default=False )
        parser.add_option( "-n", "--dry-run", dest="write_files", action="store_false", default=True )
        parser.add_option( "-v", "--verbose", dest="verbose", action="store_true", default=False )
        parser.add_option( "-d", "--debug", dest="debug", action="store_true", default=False )

        options, dummy_args = parser.parse_args( )

        self.year = int( options.year )
        self.type = options.type
        self.input = options.input
        self.output = options.output
        self.fullreport = options.fullreport
        self.write_files = options.write_files
        self.verbose = options.verbose
        self.debug = options.debug

        if self.debug:
            logging.root.setLevel( logging.DEBUG )
        elif self.verbose:
            logging.root.setLevel( logging.INFO )
        else:
            logging.root.setLevel( logging.WARN )

    def load_tourney_codes(self):
        """
        Load all of the tourney codes (and alternate names) from their data file.
        """

        LOGGER.debug( 'Loading tourney event codes' )

        codefile = csv.DictReader( open( self.EVENTCODES ) )
        for row in codefile:
            c = row['Code'].strip( )
            n = row['Name'].strip( )
            self.codes[n] = c
            self.names[c] = n
            self.tourneys.append( c )

            if row['Duration']:
                self.durations[c] = int( row['Duration'] )

            if row['Grognard']:
                self.grognards[c] = int( row['Grognard'] )

            if row['PlayLate']:
                self.playlate[c] = row['PlayLate'].strip( ).lower( )

            for altname in ['Alt1', 'Alt2', 'Alt3', 'Alt4', 'Alt5', 'Alt6']:
                if row[altname]:
                    a = row[altname].strip( )
                    self.codes[a] = c

    def load_other_codes(self):
        """
        Load all of the non-tourney codes from their data file.
        """

        LOGGER.debug( 'Loading non-tourney event codes' )

        codefile = csv.DictReader( open( self.OTHERCODES ) )
        for row in codefile:
            c = row['Code'].strip( )
            d = row['Description'].strip( )
            n = row['Name'].strip( )
            f = row['Format'].strip( )

            other = {'code': c, 'description': d, 'name': n, 'format': f}
            self.others.append( other )
            self.special.append( c )
            self.names[c] = d

    def load_preview_index(self):
        """
        Load all of the links to the event previews and map them to event codes
        """

        LOGGER.debug( 'Loading event preview index' )

        index = parse_url( self.PREVIEW_INDEX_URL % (self.year,) )
        if not index:
            LOGGER.error( 'Unable to load Preview index' )
            return

        # Find the preview table
        table = index.find( 'table' ).find( 'table' ).find( 'table' ).find( 'table' )
        rows = list( table.findAll( 'tr' ) )
        for line in range(0, len(rows), 3): # 3 rows per actual line
            top = list( rows[line].findAll('td'))
            mid = list( rows[line+1].findAll('td'))
            for column in range(0,8): # Always 8 cells
                link = name = code = None

                if line == 3 and column == 7:
                    pass

                try:
                    if type( top[column].contents[0] ) == NavigableString:
                        continue # No link, no useful data
                    link = top[column].a['href']

                    if type( mid[column].contents[0] ) == NavigableString:
                        name = normalize( unicode( mid[column].contents[0] ) ).strip( )
                        if name == 'Junior Events':
                            continue
                        m = re.match( "(.*) \((...)\)", name )
                        if m:
                            name = normalize( unicode( m.group(1) ) )
                            code = normalize( unicode( m.group(2) ) )
                        else:
                            code = normalize( unicode( mid[column].contents[2] ) ).strip( )
                    else:
                        name = normalize( unicode( mid[column].p.contents[0] ) ).strip( )
                        code = normalize( unicode( mid[column].p.contents[2] ) ).strip()

                    code = code.replace('(','').replace(')','')

                    # Map page codes to event codes
                    code = self.MISCODES[code] if self.MISCODES.has_key( code ) else code.upper()

                    self.url[code] = self.SITE_URL + link

                except Exception as e:
                    LOGGER.debug( "Skipping line %d, column %d: %s" % (line, column, e.message) )

    def check_date(self, event_date):
        """Check to see if this event date is the earliest event date seen so far"""
        if not self.first_day:
            self.first_day = event_date
        elif event_date < self.first_day:
            self.first_day = event_date
