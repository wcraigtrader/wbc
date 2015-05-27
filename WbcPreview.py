from bs4 import BeautifulSoup, Tag, NavigableString, Comment
from datetime import timedelta
from optparse import OptionParser
import logging
import codecs
import os
import re
import unicodedata
import urllib2
import xlrd

from WbcMetaData import TZ, UTC
from WbcUtility import parse_url

LOGGER = logging.getLogger( 'WbcPreview' )

DEBUGGING = True
TRAPPING = False

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
        cls.INITIALIZED = True

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
        cls.LOOKUP[ 'Cromwell' ] = cls.LOOKUP[ 'Cornwall' ]
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

        cls.LOOKUP[ 'HMWG' ] = Token( 'Format', 'HMWG' )

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
            'sty_cont' : Token.CONTINUOUS,
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
            'for_mese' : Token( 'Format', 'Heats' ),
            'for_se' : Token( 'Format', 'Single Elimination' ),
            'for_sem' : Token( 'Format', 'Single Elimination Mulligan' ),
            'for_swis' : Token( 'Format', 'Swiss' ),
            'for_swel' : Token( 'Format', 'Swiss Elimination' ),
            'freeform' : Token( 'Style', 'Freeform' ),
            'sty_sche' : Token( 'Style', 'Scheduled' ),
            'sty_ctht' : Token( 'Style', 'Continuous Heats' ),
            'prize1' : Token( 'Prizes', 1, 1 ),
            'prize2' : Token( 'Prizes', 2, 2 ),
            'prize3' : Token( 'Prizes', 3, 3 ),
            'prize4' : Token( 'Prizes', 4, 4 ),
            'prize5' : Token( 'Prizes', 5, 5 ),
            'prize6' : Token( 'Prizes', 6, 6 ),
            'prztrial' : Token( 'Prizes', 'Trial', 0 ),
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
            elif isinstance( tag, Tag ) and tag.name in [ 'img' ]:
                tokens += cls.tokenize_text( buffer )
                buffer = u''
                tokens += cls.tokenize_icon( tag )
            else:
                pass  # ignore other tags, for now
                # LOGGER.debug( 'Ignored <%s>', tag.name )

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
        elif name in [ 'stadium', 'class_a', 'class_b', 'coached' ]:
            pass
        elif name.startswith( 'for_' ):
            format = name[4:]
            token = Token( 'Format', format )
            cls.ICONS[ name ] = token
            tokens.append( token )
            LOGGER.warn( 'Automatically added format [%s]', format )
        elif name.startswith( 'sty_' ):
            style = name[4:]
            token = Token( 'Style', style )
            cls.ICONS[ name ] = token
            tokens.append( token )
            LOGGER.warn( 'Automatically added style [%s]', style )
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
                    m = re.match( "[A-Z0-9&]{3}", data )
                    if m:
                        n = m.group()
                        tokens.append( Token( 'Code', n ) )
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
            LOGGER.debug( 'Skipped [%s] in [%s]', hjunk, hdata )

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

    def match_one_or_more_days( self, pos=0 ):
        self.match_initialize()

        if not self.have( pos, 'Day' ):
            return 0

        l = 0
        while self.have( pos + l, 'Day' ):
            self.last_day = self.tokens[pos]
            l += 1

        LOGGER.debug( 'Matched %s => %s', self.tokens[pos:pos + 1], self.last_day )
        del self.tokens[pos:pos + l]
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

            if self.have( pos + l, Token.CONTINUOUS ):
                self.last_continuous = True
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
            return 0

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

        if not self.have( pos, 'Room' ):
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
    # codemap = { 'kot' : 'KOT' }
    codemap = { 'gmb' : 'GBM' }

    # TODO: Preview codes for messages
    notes = {
        'CNS': "Can't match 30 minute rounds",
        'ELC': "Can't match 20 minute rounds",
        'LID': "Can't match 20 minute rounds",
        'SLS': "Can't match 20 minute rounds",
        'LST': "Can't match 30 minute rounds, Can't handle 'until conclusion'",
        'KOT': "Can't match 30 minute rounds, Can't handle 'to conclusion'",
        'PGF': "Can't handle 'to conclusion'",

#         'ADV': "Preview shows 1 hour for SF and F, not 2 hours per spreadsheet and pocket schedule",
#         'BAR': "Pocket schedule shows R1@8/8:9, R2@8/8:14, R3@8/8:19, SF@8/9:9, F@8/9:14",
#         'KFE': "Pocket schedule shows R1@8/6:9, R2@8/6:16, SF@8/7:9, F@8/7:16",
#         'MED': "Preview shows heats taking place after SF/F",
#         'RFG': "Should only have 1 demo, can't parse H1 room",
#         'STA': "Conestoga is misspelled as Coonestoga",
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
        code_tokens = None
        event_map = None
        events = None

        def __init__( self, code, name, page, first_day ):
            """The schedule table within the page is a table that has a variable number of rows:

            [0] Contains the date the page was last updated -- ignored (may not be present).
            [1] Contains the token code and other image codes
            [2:-2] Contains the schedule data, mostly as images, in two columns
            [-1] Contains the location information.
            """
            self.code = code
            self.name = name
            self.first_day = first_day

            if self.code == 'ACQ':
                pass

            # Find schedule / rows
            tables = page.findAll( 'table' )
            schedule = tables[2]
            rows = schedule.findAll( 'tr' )

            if rows[0].text.find( 'Updated' ) != -1:
                rows = rows[1:]  # Remove optional and ignored update time

            self.tokenize_codes( rows )
            self.tokenize_rooms( rows )
            self.tokenize_times( rows )

            if DEBUGGING:
                LOGGER.info( "%3s: rooms %s", self.code, Token.dump_list( self.room_tokens ) )
                LOGGER.info( "     times %s", Token.dump_list( self.event_tokens ) )
                LOGGER.info( "     codes %s", Token.dump_list( self.code_tokens ) )

            self.parse_rooms()
            self.parse_events()

        def tokenize_codes( self, rows ):
            self.code_tokens = Token.tokenize( rows[0] )

        def tokenize_times( self, rows ):
            """Parse 3rd row through next-to-last for event data"""

            self.event_tokens = []
            start = 1
            for row in rows[start:-1]:
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

        def parse_events_2015( self ):

            # START <day> ( <event> <time> { PLUS | MINUS <end time> }? ( AT <room> )? )*
            # START <day>* ( <day> <event> <time> PLUS? )? <day>* <event> <time> PLUS?
            # [ AT <room> ] }+
            # If there are multiple times, then each represents a separate event on that day

            LOGGER.debug( 'Parsing events ...' )

            awards_are_events = self.code in ['UPF', 'VIP', 'WWR' ]
            two_weekends = self.code in [ 'AFK', 'AOR', 'AUC', 'BWD', 'GBG', 'PZB', 'TRC', 'SQL', 'WAT', 'WSM' ]

            self.events = []

            weekend_offset = 0
            weekend_start = 5 if two_weekends else 1

            p = Parser( self.event_tokens )

            if self.code in [ 'ATS']:
                pass

            while p.have_start():
                p.match( Token.START )

                while p.count and not p.have_start():
                    if p.match_day():
                        day = p.last_day
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
                        if self.code == 'PDT' and p.last_name.endswith( 'FC' ):
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, self.draft_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name + u' Draft', TZ.localize( dtime ), room )
                            self.events.append ( e )

                            dtime = dtime + timedelta( hours=1 )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, TZ.localize( dtime ), self.default_room )
                            self.events.append ( e )

                        else:
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, p.last_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, TZ.localize( dtime ), room )
                            self.events.append ( e )

                    else:
                        p.recover()

            self.events.sort()

        def parse_events( self ):

            # START <day> ( <event> <time> { PLUS | MINUS <end time> }? ( AT <room> )? )*
            # START <day> ( <day>* <event> <time> PLUS? )? <day>* <event> <time> PLUS?
            # [ AT <room> ] }+
            # If there are multiple days, then all of the events happen on all of the days (except demos)
            # If there are multiple times, then each represents a separate event on that day

            LOGGER.debug( 'Parsing events ...' )

            awards_are_events = self.code in ['UPF', 'VIP', 'WWR' ]
            two_weekends = self.code in [ 'AFK', 'AOR', 'AUC', 'BWD', 'GBG', 'PZB', 'TRC', 'SQL', 'WAT', 'WSM' ]

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
                    if p.match_one_or_more_days():
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
        """Load and parse the preview page for a single tournament.
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

