# ----- Copyright (c) 2010-2016 by W. Craig Trader ---------------------------------
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

from bs4 import Tag, NavigableString, Comment
from datetime import timedelta
import logging
import re

from WbcUtility import parse_url, localize

LOG = logging.getLogger( 'WbcPreview' )


# ----- Token -----------------------------------------------------------------

class Token( object ):
    """Simple data object for breaking descriptions into parseable tokens"""

    INITIALIZED = False

    LOOKUP = {}
    DAYS = {}
    ICONS = {}

    def __init__(self, t, l=None, v=None):
        self.type = t
        self.label = l
        self.value = v

    def __str__(self):
        return str( self.label ) if self.label else self.type

    def __repr__(self):
        return self.__str__( )

    def __eq__(self, other):
        return self.type == other.type and self.label == other.label and self.value == other.value

    def __ne__(self, other):
        return not self.__eq__( other )

    @classmethod
    def initialize(cls):
        if cls.INITIALIZED: return
        cls.INITIALIZED = True

        cls.AM = cls.LOOKUP['AM'] = Token( 'Meridian', 'AM' )
        cls.PM = cls.LOOKUP['PM'] = Token( 'Meridian', 'PM' )

        cls.add_day( 'SAT1', 0, 'First Saturday' )
        cls.add_day( 'SUN1', 1, 'First Sunday' )
        cls.add_day( 'MON', 2, 'Monday' )
        cls.add_day( 'TUE', 3, 'Tuesday' )
        cls.add_day( 'WED', 4, 'Wednesday' )
        cls.add_day( 'THU', 5, 'Thursday' )
        cls.add_day( 'FRI', 6, 'Friday', 'Fr' )
        cls.add_day( 'SAT2', 7, 'Saturday' )
        cls.add_day( 'SUN2', 8, 'Sunday' )

        cls.add_event( 'Demo', 'Demonstration', 'Demonstrations' )
        cls.add_event( 'H', 'Heat' )
        cls.add_event( 'R', 'Round' )
        cls.add_event( 'QF', 'Quarterfinal', 'Quarterfinals' )
        cls.add_event( 'SF', 'Semifinal', 'Semifinals' )
        cls.add_event( 'F', 'Final', 'Finals' )
        cls.add_event( 'Junior' )
        cls.add_event( 'Mulligan', 'mulligan' )
        cls.add_event( 'After Action', 'After Action Briefing', 'After Action Meeting' )
        cls.add_event( 'Seminar', 'Optional Seminar' )
        cls.add_event( 'Draft', 'DRAFT' )

        # cls.add_event( 'H1' )
        # cls.add_event( 'H2' )
        # cls.add_event( 'H3' )
        # cls.add_event( 'H4' )
        # cls.add_event( 'R1' )
        # cls.add_event( 'R2' )
        # cls.add_event( 'R3' )
        # cls.add_event( 'R4' )
        # cls.add_event( 'R5' )
        # cls.add_event( 'R6' )

        cls.add_qualifier( 'PC', 'Grognard PC' )
        cls.add_qualifier( 'AFC' )
        cls.add_qualifier( 'NFC' )
        cls.add_qualifier( 'Super Bowl' )

        cls.CONTINUOUS = Token( 'Symbol', '...' )
        cls.LOOKUP['to completion'] = Token.CONTINUOUS
        cls.LOOKUP['till completion'] = Token.CONTINUOUS
        cls.LOOKUP['until completion'] = Token.CONTINUOUS
        cls.LOOKUP['until conclusion'] = Token.CONTINUOUS
        cls.LOOKUP['to conclusion'] = Token.CONTINUOUS

        cls.SHIFT = Token( 'Symbol', '>' )
        cls.LOOKUP['moves to'] = Token.SHIFT
        cls.LOOKUP['moving to'] = Token.SHIFT
        cls.LOOKUP['shifts to'] = Token.SHIFT
        cls.LOOKUP['switches to'] = Token.SHIFT
        cls.LOOKUP['switching to'] = Token.SHIFT
        cls.LOOKUP['after drafts in'] = Token.SHIFT

        cls.LOOKUP['HMWG'] = Token( 'Format', 'HMWG' )

        cls.initialize_7springs_rooms( )

        cls.PATTERN = '|'.join( sorted( cls.LOOKUP.keys( ), reverse=True ) )

        cls.AT = cls.LOOKUP['@'] = Token( 'Symbol', '@' )
        cls.AND = cls.LOOKUP['&'] = Token( 'Symbol', '&' )
        cls.PLUS = cls.LOOKUP['+'] = Token( 'Symbol', '+' )
        cls.DASH = cls.LOOKUP['-'] = Token( 'Symbol', '-' )
        cls.SLASH = cls.LOOKUP['/'] = Token( 'Symbol', '/' )

        cls.PATTERN += '|[@&+-/]'

        cls.START = cls.LOOKUP['|'] = Token( 'Symbol', '|' )

    @classmethod
    def add_event(cls, primary, *aliases):
        room = Token( 'Event', primary )
        cls.LOOKUP[primary] = room
        for alias in aliases:
            cls.LOOKUP[alias] = room

    @classmethod
    def add_qualifier(cls, primary, *aliases):
        room = Token( 'Qualifier', primary )
        cls.LOOKUP[primary] = room
        for alias in aliases:
            cls.LOOKUP[alias] = room

    @classmethod
    def add_day(cls, name, value, *aliases):
        day = Token( 'Day', name, value )
        cls.DAYS[name] = day
        cls.LOOKUP[name] = day
        for alias in aliases:
            cls.LOOKUP[alias] = day

    @classmethod
    def add_room(cls, primary, *aliases):
        room = Token( 'Room', primary )
        cls.LOOKUP[primary] = room
        for alias in aliases:
            cls.LOOKUP[alias] = room

    @classmethod
    def initialize_7springs_rooms(cls):
        cls.add_room( 'Alpine' )
        cls.add_room( 'Ballroom', 'Ballroom B' )
        cls.add_room( 'Ballroom Stage' )
        cls.add_room( 'Bavarian Lounge' )
        cls.add_room( 'Chestnut' )
        cls.add_room( 'Dogwood', 'Dogwood Forum' )
        cls.add_room( 'Evergreen' )
        cls.add_room( 'Evergreen & Chestnut' )
        cls.add_room( 'Exhibit Annex T1', 'Exhibit Annex Table 1', 'Exhibit annex Table 1' )
        cls.add_room( 'Exhibit Annex T2', 'Exhibit Annex Table 2', 'Exhibit annex Table 2' )
        cls.add_room( 'Exhibit Annex T3', 'Exhibit Annex Table 3', 'Exhibit annex Table 3' )
        cls.add_room( 'Exhibit Annex T4', 'Exhibit Annex Table 4', 'Exhibit annex Table 4' )
        cls.add_room( 'Exhibit Annex T5', 'Exhibit Annex Table 5', 'Exhibit annex Table 5' )
        cls.add_room( 'Exhibit Annex T6', 'Exhibit Annex Table 6', 'Exhibit annex Table 6' )
        cls.add_room( 'Exhibit Annex T7', 'Exhibit Annex Table 7', 'Exhibit annex Table 7' )
        cls.add_room( 'Exhibit Annex T8', 'Exhibit Annex Table 8', 'Exhibit annex Table 8' )
        cls.add_room( 'Exhibit Annex T9', 'Exhibit Annex Table 9', 'Exhibit annex Table 9' )
        cls.add_room( 'Exhibit Hall' )
        cls.add_room( 'Festival Hall', 'Festival' )
        cls.add_room( 'First Tracks Center', 'Ski Lodge First Tracks Center', 'Ski Lodge Fast Tracks Center' )
        cls.add_room( 'First Tracks Poolside', 'Ski Lodge First Tracks Poolside' )
        cls.add_room( 'First Tracks Slopeside', 'Ski Lodge First Tracks Slopeside' )
        cls.add_room( 'Foggy Goggle Center', 'Ski Lodge Foggy Goggle Center' )
        cls.add_room( 'Foggy Goggle Front', 'Ski Lodge Foggy Goggle Front', 'Ski Lodge Foggie Goggle Front' )
        cls.add_room( 'Foggy Goggle Rear', 'Ski Lodge Foggy Goggle Rear' )
        cls.add_room( 'Fox Den' )
        cls.add_room( 'Hemlock' )
        cls.add_room( 'Laurel' )
        cls.add_room( 'Maple Room', 'Ski Lodge Maple Room' )
        cls.add_room( 'Rathskeller' )
        cls.add_room( 'Seasons' )
        cls.add_room( 'Seasons 1' )
        cls.add_room( 'Seasons 1-2' )
        cls.add_room( 'Seasons 1-3' )
        cls.add_room( 'Seasons 1-4' )
        cls.add_room( 'Seasons 1-5' )
        cls.add_room( 'Seasons 2' )
        cls.add_room( 'Seasons 2-4' )
        cls.add_room( 'Seasons 2-5' )
        cls.add_room( 'Seasons 3' )
        cls.add_room( 'Seasons 3-4' )
        cls.add_room( 'Seasons 3-5' )
        cls.add_room( 'Seasons 4-5' )
        cls.add_room( 'Seasons 5' )
        cls.add_room( 'Snowflake Forum' )
        cls.add_room( 'Stag Pass' )
        cls.add_room( 'Sunburst Forum' )
        cls.add_room( 'Timberstone' )
        cls.add_room( 'Wintergreen' )

    @classmethod
    def initialize_host_rooms(cls):
        cls.LOOKUP['Ballroom A'] = Token( 'Room', 'Ballroom A' )
        cls.LOOKUP['Ballroom B'] = Token( 'Room', 'Ballroom B' )
        cls.LOOKUP['Ballroom AB'] = Token( 'Room', 'Ballroom AB' )
        cls.LOOKUP['Ballroom'] = cls.LOOKUP['Ballroom AB']

        cls.LOOKUP['Conestoga 1'] = Token( 'Room', 'Conestoga 1' )
        cls.LOOKUP['Conestoga 2'] = Token( 'Room', 'Conestoga 2' )
        cls.LOOKUP['Conestoga 3'] = Token( 'Room', 'Conestoga 3' )
        cls.LOOKUP['Coonestoga 3'] = cls.LOOKUP['Conestoga 3']

        cls.LOOKUP['Cornwall'] = Token( 'Room', 'Cornwall' )
        cls.LOOKUP['Cromwell'] = cls.LOOKUP['Cornwall']
        cls.LOOKUP['Heritage'] = Token( 'Room', 'Heritage' )
        cls.LOOKUP['Hopewell'] = Token( 'Room', 'Hopewell' )
        cls.LOOKUP['Kinderhook'] = Token( 'Room', 'Kinderhook' )
        cls.LOOKUP['Lampeter'] = Token( 'Room', 'Lampeter' )
        cls.LOOKUP['Laurel Grove'] = Token( 'Room', 'Laurel Grove' )
        cls.LOOKUP['Limerock'] = Token( 'Room', 'Limerock' )
        cls.LOOKUP['Marietta'] = Token( 'Room', 'Marietta' )
        cls.LOOKUP['New Holland'] = Token( 'Room', 'New Holland' )
        cls.LOOKUP['Paradise'] = Token( 'Room', 'Paradise' )
        cls.LOOKUP['Showroom'] = Token( 'Room', 'Showroom' )
        cls.LOOKUP['Strasburg'] = Token( 'Room', 'Strasburg' )
        cls.LOOKUP['Wheatland'] = Token( 'Room', 'Wheatland' )

        cls.LOOKUP['Terrace 1'] = Token( 'Room', 'Terrace 1' )
        cls.LOOKUP['Terrace 2'] = Token( 'Room', 'Terrace 2' )
        cls.LOOKUP['Terrace 3'] = Token( 'Room', 'Terrace 3' )
        cls.LOOKUP['Terrace 4'] = Token( 'Room', 'Terrace 4' )
        cls.LOOKUP['Terrace 5'] = Token( 'Room', 'Terrace 5' )
        cls.LOOKUP['Terrace 6'] = Token( 'Room', 'Terrace 6' )
        cls.LOOKUP['Terrace 7'] = Token( 'Room', 'Terrace 7' )

        cls.LOOKUP['Vista C'] = Token( 'Room', 'Vista C' )
        cls.LOOKUP['Vista D'] = Token( 'Room', 'Vista D' )
        cls.LOOKUP['Vista CD'] = Token( 'Room', 'Vista CD' )
        cls.LOOKUP['Vista'] = cls.LOOKUP['Vista CD']

    @classmethod
    def initialize_icons(cls):
        cls.ICONS = {
            'semi': Token( 'Award', 'SF' ),
            'final': Token( 'Award', 'F' ),
            'heat1': cls.LOOKUP['H1'],
            'heat2': cls.LOOKUP['H2'],
            'heat3': cls.LOOKUP['H3'],
            'heat4': cls.LOOKUP['H4'],
            'rd1': cls.LOOKUP['R1'],
            'rd2': cls.LOOKUP['R2'],
            'rd3': cls.LOOKUP['R3'],
            'rd4': cls.LOOKUP['R4'],
            'rd5': cls.LOOKUP['R5'],
            'rd6': cls.LOOKUP['R6'],
            'sty_cont': Token.CONTINUOUS,
            'demo': cls.LOOKUP['Demo'],
            'demoweb': cls.LOOKUP['Demo'],
            'demo_folder_transparent': cls.LOOKUP['Demo'],
            'jrwebicn': cls.LOOKUP['Junior'],
            'mulligan': cls.LOOKUP['Mulligan'],
            'sat': cls.DAYS['SAT'],
            'sun': cls.DAYS['SUN'],
            'mon': cls.DAYS['MON'],
            'tue': cls.DAYS['TUE'],
            'wed': cls.DAYS['WED'],
            'thu': cls.DAYS['THU'],
            'fri': cls.DAYS['FRI'],
            'sat2': cls.DAYS['SAT'],
            'sun2': cls.DAYS['SUN'],
            'mon2': cls.DAYS['MON'],
            'tue2': cls.DAYS['TUE'],
            'wed2': cls.DAYS['WED'],
            'thu2': cls.DAYS['THU'],
            'fri2': cls.DAYS['FRI'],
            'for_mese': Token( 'Format', 'Heats' ),
            'for_se': Token( 'Format', 'Single Elimination' ),
            'for_sem': Token( 'Format', 'Single Elimination Mulligan' ),
            'for_swis': Token( 'Format', 'Swiss' ),
            'for_swel': Token( 'Format', 'Swiss Elimination' ),
            'freeform': Token( 'Style', 'Freeform' ),
            'sty_sche': Token( 'Style', 'Scheduled' ),
            'sty_ctht': Token( 'Style', 'Continuous Heats' ),
            'prize1': Token( 'Prizes', 1, 1 ),
            'prize2': Token( 'Prizes', 2, 2 ),
            'prize3': Token( 'Prizes', 3, 3 ),
            'prize4': Token( 'Prizes', 4, 4 ),
            'prize5': Token( 'Prizes', 5, 5 ),
            'prize6': Token( 'Prizes', 6, 6 ),
            'prztrial': Token( 'Prizes', 'Trial', 0 ),
        }

    @staticmethod
    def encode_list(tokenlist):
        return [u' '.join( [unicode( y ) for y in x] ) for x in tokenlist]

    @classmethod
    def tokenize(cls, tag):
        tokens = []
        partial = u''

        for tag in tag.descendants:
            if isinstance( tag, Comment ):
                pass  # Always ignore comments
            elif isinstance( tag, NavigableString ):
                partial += u' ' + unicode( tag )
            elif isinstance( tag, Tag ) and tag.name in ['img']:
                tokens += cls.tokenize_text( partial )
                partial = u''
                tokens += cls.tokenize_icon( tag )
            else:
                pass  # ignore other tags, for now
                # LOG.debug( 'Ignored <%s>', tag.name )

        if partial:
            tokens += cls.tokenize_text( partial )

        return tokens

    @classmethod
    def tokenize_icon(cls, tag):
        cls.initialize( )

        tokens = []

        try:
            name = tag['src'].lower( )
            name = name.split( '/' )[-1]
            name = name.split( '.' )[0]
        except:
            LOG.error( "%s didn't have a 'src' attribute", tag )
            return tokens

        if cls.ICONS.has_key( name ):
            tokens.append( cls.ICONS[name] )
        elif name in ['stadium', 'class_a', 'class_b', 'coached']:
            pass
        elif name.startswith( 'for_' ):
            form = name[4:]
            token = Token( 'Format', form )
            cls.ICONS[name] = token
            tokens.append( token )
            LOG.warn( 'Automatically added form [%s]', form )
        elif name.startswith( 'sty_' ):
            style = name[4:]
            token = Token( 'Style', style )
            cls.ICONS[name] = token
            tokens.append( token )
            LOG.warn( 'Automatically added style [%s]', style )
        else:
            LOG.warn( 'Ignored icon [%s]', name )

        return tokens

    @classmethod
    def tokenize_text(cls, text):
        cls.initialize( )

        data = text

        junk = u''
        tokens = []

        # Cleanup crappy data
        data = data.replace( u'\xa0', u' ' )
        data = data.replace( u'\n', u' ' )

        data = data.strip( )
        data = data.replace( u' ' * 11, u' ' ).replace( u' ' * 7, u' ' ).replace( u' ' * 5, u' ' )
        data = data.replace( u' ' * 3, u' ' ).replace( u' ' * 2, u' ' )
        data = data.replace( u' ' * 2, u' ' ).replace( u' ' * 2, u' ' )
        data = data.strip( )

        hdata = data.encode( 'unicode_escape' )

        while len( data ):

            # Ignore commas and semi-colons
            if data[0] in u',;:':
                data = data[1:]
            else:
                # Match Room names, event names, phrases, symbols
                m = re.match( cls.PATTERN, data )
                if m:
                    text = m.group( )
                    tokens.append( cls.LOOKUP[text] )
                    data = data[len( text ):]
                else:
                    # Match numbers
                    m = re.match( "\d+", data )
                    if m:
                        text = m.group( )
                        n = int( text )
                        tokens.append( Token( 'Number', text, n ) )
                        data = data[len( text ):]
                    else:
                        junk += data[0];
                        data = data[1:]

            data = data.strip( )

        if junk:
            hjunk = junk.encode( 'unicode_escape' )
            LOG.debug( 'Skipped [%s] in [%s]', hjunk, hdata )

        return tokens


class Parser( object ):
    tokens = []
    last_match = None

    def __init__(self, tokens):
        self.tokens = list( tokens )

    def __str__(self):
        return "%s (%d) %s" % (self.last_match, self.count, self.tokens)

    @property
    def count(self):
        return len( self.tokens )

    def has(self, *prediction):
        lookahead_length = len( prediction )
        if self.count < lookahead_length:
            return False

        pos = 0
        for predicted in prediction:
            lookahead = self.tokens[pos]
            pos += 1

            if isinstance( predicted, Token ):
                if lookahead == predicted:
                    continue
            elif lookahead.type == predicted:
                continue

            return False

        return True

    def is_not(self, *stops):
        if self.count < 1:
            return True

        current = self.tokens[0]
        for stop in stops:
            if isinstance( stop, Token ):
                if current == stop:
                    return False
            elif current.type == stop:
                return False
        return True

    def next(self):
        return self.tokens.pop( 0 )

    def match_3_events(self):
        """
        Match 'Event', Token.SLASH, 'Event', Token.SLASH, 'Event'
        Stop [ 'Day' ]
        """
        e1 = self.next( )
        self.next( )
        e2 = self.next( )
        self.next( )
        e3 = self.next( )
        return [(e1.label,), (e2.label,), (e3.label,)]

    def match_2_events(self):
        """
        Match 'Event', Token.SLASH, 'Event'
        Stop [ 'Day' ]
        """
        e1 = self.next( )
        self.next( )
        e2 = self.next( )
        return [(e1.label,), (e2.label,)]

    def match_event(self):
        """
        Match 'Event'
        Stop [ 'Day' ]
        """
        e1 = self.next( )
        return [(e1.label,)]

    def match_multiple_heats(self):
        """
        Match 'Event', 'Number', Token.DASH, 'Number'
        Stop [ 'Day' ]
        """
        e1 = self.next( )
        first = self.next( )
        self.next( )
        last = self.next( )

        results = []
        for n in range( first.value, last.value + 1 ):
            results.append( (e1.label, n) )

        return results

    def match_heat(self):
        """
        Match 'Event', 'Number', 'Qualifier'
        Match 'Event', 'Number'
        Stop [ 'Day' ]
        """
        e1 = self.next( )
        n = self.next( )
        if self.has( 'Qualifier' ):
            q = self.next( )
            return [(e1.label, n.value, q.label)]

        return [(e1.label, n)]

    def match_round(self):
        """
        Match 'Event', 'Number', Token.SLASH, 'Number'
        Stop [ 'Day' ]
        """
        e1 = self.next( )
        n = self.next( )
        self.next( )
        m = self.next( )

        return [(e1.label, n.value, m.value)]

    def match_date_times(self):
        """Recognized date/time formats:
    
        <day> <time> @
        <day> @ <time>
        <day> <time> & <day> <time> @
        <day> <time> <time> <time> & <time> @
        <day> <time> & <day> <time> & <time> & <day> <time> & <day> <time> @
        <day> @ <time> <time> <time> <time>
    
        Ignores [ &, @ ]
        Stops on [ -?,  <room> ]
    
        Returns a list of ( Day, Time ) tuples
        """

        results = []
        while self.has( 'Day' ):
            day = self.next( )
            while self.is_not( Token.DASH, 'Room', 'Day' ):
                if self.has( 'Number' ):
                    offset = 0
                    time = self.next( )
                    if self.has( Token.AM ):
                        self.next( )
                    elif self.has( Token.PM ):
                        self.next( )
                        offset = 12
                    elif self.has( Token.PLUS ):
                        self.next( )  # FIXME: Do something with + => continuous???
                    results.append( timedelta( days=day.value, hours=time.value + offset ) )
                elif self.has( Token.AND ) or self.has( Token.AT ):
                    self.next( )

        return results

    def match_room(self):
        if self.count and self.tokens[0] == Token.DASH:
            self.next( )

        if self.count and self.tokens[0].type == 'Room':
            return self.next( )

        return None


# ----- WBC Preview Schedule -------------------------------------------------

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

    # TODO: Preview codes for messages, move to Meta
    notes = {
        'CNS': "Can't match 15 minute rounds",
        'ELC': "Can't match 20 minute rounds",
        'KOT': "Can't match 15 minute rounds, Can't handle 'to conclusion'",
        'LID': "Can't match 30 minute rounds",
        'LST': "Can't match 30 minute rounds, Can't handle 'until conclusion'",
        'PGF': "Can't handle 'to conclusion'",
        'RTT': "Can't match 30 minute rounds, Can't handle 'until completion'",
        'SLS': "Can't match 20 minute rounds",

        # 'ADV': "Preview shows 1 hour for SF and F, not 2 hours per spreadsheet and pocket schedule",
        # 'BAR': "Pocket schedule shows R1@8/8:9, R2@8/8:14, R3@8/8:19, SF@8/9:9, F@8/9:14",
        # 'KFE': "Pocket schedule shows R1@8/6:9, R2@8/6:16, SF@8/7:9, F@8/7:16",
        # 'MED': "Preview shows heats taking place after SF/F",
        # 'RFG': "Should only have 1 demo, can't parse H1 room",
        # 'STA': "Conestoga is misspelled as Coonestoga",
    }

    class Event( object ):
        """Simple data object to collect information about an event occuring at a specific time."""

        code = None
        name = None
        type = None
        time = None
        location = None

        def __init__(self, code, name, etype, etime, location):
            self.code = code
            self.name = name
            self.type = etype
            self.time = etime
            self.location = location

        def __cmp__(self, other):
            return cmp( self.time, other.time )

        def __str__(self):
            return '%s %s %s in %s at %s' % (self.code, self.name, self.type, self.location, self.time)

    class Tourney( object ):

        tracking = []

        def __init__(self, meta, code, name, page):

            self.meta = meta
            self.code = code
            self.name = name

            self.events = []
            self.event_tokens = []

            td = page.table.table.table.findAll( 'tr' )[2].td
            paras = list( td.findAll( 'p' ) )

            if len( paras ):
                self.tokenize_events( paras )
                self.parse_events( )

            else:
                LOG.error( '%s: Did not find schedule data', code )

        def tokenize_events(self, paras):
            for para in paras:
                text = para.text.strip( )
                if text.startswith( 'Demo' ) and self.code not in ['775', '7WD', '989', 'AOA']:
                    self.event_tokens.append( Token.tokenize_text( text ) )
                elif self.code in ['HRC']:
                    if text:
                        self.event_tokens.append( Token.tokenize_text( text ) )
                else:
                    for line in text.split( '\n' ):
                        stripped = line.strip( )
                        if stripped:
                            self.event_tokens.append( Token.tokenize_text( stripped ) )

            # if self.code in self.tracking:
            for section in Token.encode_list( self.event_tokens ):
                LOG.debug( "%s: %s", self.code, section )

        def parse_events(self):
            """Recognized event formats:

            Demo <day> <time> @ <room>
            Demo <day> @ <time> -? <room>
            Demo <day> <time> & <day> <time> @ <room>
            Demo <day> <time> <time> <time> & <time> @ <room>
            Demo <day> <time> & <day> <time> & <time> & <day> <time> & <day> <time> @ <room>
            <event> <day> @ <time> -? <room>
            <event> # <day> @ <time> -? <room>
            <event> # <qualifier> <day> @ <time> -? <room>
            <event> # / # <day> @ <time> -? <room>
            <event> # - # <day> @ <time> <time> <time> <time> -? <room>
            QF / SF / F <day> @ <time> -? <room>
            QF / SF <day> @ <time> -? <room>
            SF / F <day> @ <time> -? <room>
            """

            # return

            for row in self.event_tokens:
                p = Parser( row )
                if p.has( 'Event', Token.SLASH, 'Event', Token.SLASH, 'Event', 'Day' ):
                    elist = p.match_3_events( )
                elif p.has( 'Event', Token.SLASH, 'Event', 'Day' ):
                    elist = p.match_2_events( )
                elif p.has( 'Event', 'Number', Token.DASH, 'Number', 'Day' ):
                    elist = p.match_multiple_heats( )
                elif p.has( 'Event', 'Number', Token.SLASH, 'Number', 'Day' ):
                    elist = p.match_round( )
                elif p.has( 'Event', 'Number', 'Qualifier', 'Day' ):
                    elist = p.match_heat( )
                elif p.has( 'Event', 'Number', 'Day' ):
                    elist = p.match_heat( )
                elif p.has( 'Event', 'Day' ):
                    elist = p.match_event( )
                else:
                    LOG.error( '%s: Could not match %s', self.code, row )
                    continue

                times = p.match_date_times( )
                room = p.match_room( )

                # self.add_events( elist, times, room )

            self.events.sort( )

        def add_events(self, elist, times, room):
            if len( elist ) == len( times ):
                for e, t in zip( elist, times ):
                    self.events.append(
                        WbcPreview.Event( self.code, self.name, e[0], localize( self.meta.first_day + t ), room ) )

    class Pre2016Tourney( object ):
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

        tracking = []

        def __init__(self, meta, code, name, page):
            """The schedule table within the page is a table that has a variable number of rows:

            [0] Contains the date the page was last updated -- ignored (may not be present).
            [1] Contains the token code and other image codes
            [2:-2] Contains the schedule data, mostly as images, in two columns
            [-1] Contains the location information.
            """

            self.meta = meta
            self.code = code
            self.name = name
            self.first_day = self.meta.first_day

            if self.code == 'ACQ':
                pass

            # Find schedule / rows
            rows = page.findAll( 'table' )[2].findAll( 'tr' )

            if rows[0].text.find( 'Updated' ) != -1:
                rows = rows[1:]  # Remove optional and ignored update time

            self.tokenize_codes( rows )
            self.tokenize_rooms( rows )
            self.tokenize_times( rows )

            if self.code in self.tracking:
                LOG.info( "%3s: rooms %s", self.code, Token.dump_list( self.room_tokens ) )
                LOG.info( "     times %s", Token.dump_list( self.event_tokens ) )
                LOG.info( "     codes %s", Token.dump_list( self.code_tokens ) )

            self.parse_rooms( )
            self.parse_events( )

        def tokenize_codes(self, rows):
            self.code_tokens = Token.tokenize( rows[0] )

        def tokenize_times(self, rows):
            """Parse 3rd row through next-to-last for event data"""

            self.event_tokens = []
            start = 1
            for row in rows[start:-1]:
                for td in row.findAll( 'td' ):
                    self.event_tokens.append( Token.START )
                    self.event_tokens += Token.tokenize( td )

        def tokenize_rooms(self, rows):
            """"parse last row for room data"""

            self.room_tokens = Token.tokenize( rows[-1] )

        def parse_rooms(self):

            # <default room>? [ SHIFT <shift room> [ AT <day> <time> ] ] { <event>+ <room> }*
            #
            # If this is PDT, then the shift room is special

            LOG.debug( 'Parsing rooms ...' )

            self.event_map = {}

            p = Parser( self.room_tokens )

            if p.match_room_shift( ):
                self.default_room = p.default_room
                if self.code == 'PDT':
                    self.draft_room = p.shift_room
                else:
                    self.shift_room = p.shift_room
                    self.shift_time = self.first_day + timedelta( days=p.shift_day.value ) + p.shift_time.value

            elif p.match_room( ):
                self.default_room = p.last_room

            while p.match_room_events( ):
                for e in p.event_list:
                    self.event_map[e] = p.last_room

            if p.count:
                LOG.warn( '%3s: Unmatched room tokens: %s', self.code, p.tokens )
                pass

        def parse_events(self):
            if self.code in ['AFK', 'PZB', 'WAW']:
                self.parse_events_2014( )
            else:
                self.parse_events_2015( )

        def parse_events_2015(self):

            # START <day> ( <event> <time> { PLUS | MINUS <end time> }? ( AT <room> )? )*
            # START <day>* ( <day> <event> <time> PLUS? )? <day>* <event> <time> PLUS?
            # [ AT <room> ] }+
            # If there are multiple times, then each represents a separate event on that day

            LOG.debug( 'Parsing events ...' )

            awards_are_events = self.code in ['UPF', 'VIP', 'WWR']
            two_weekends = self.code in ['AFK', 'AOR', 'AUC', 'BWD', 'GBG', 'PZB', 'TRC', 'SQL', 'WAT', 'WSM']

            self.events = []

            weekend_offset = 0
            weekend_start = 5 if two_weekends else 1

            p = Parser( self.event_tokens )

            if self.code in ['ATS']:
                pass

            while p.have_start( ):
                p.match( Token.START )

                while p.count and not p.have_start( ):
                    if p.match_day( ):
                        day = p.last_day
                        weekend_offset = 7 if day.value > weekend_start else weekend_offset
                        day_of_week = day.value + weekend_offset if day.value < 2 else day.value
                        midnight = self.first_day + timedelta( days=day_of_week )

                    elif p.match_multiple_event_times( ):
                        for etime in p.time_list:
                            # handle events immediately
                            dtime = midnight + etime
                            room = self.find_room( p.last_actual, dtime, p.last_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, localize( dtime ), room )
                            self.add_event( e )

                    elif p.match_single_event_time( awards_are_events=awards_are_events ):
                        if self.code == 'PDT' and p.last_name.endswith( 'FC' ):
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, self.draft_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name + u' Draft', localize( dtime ),
                                                  room )
                            self.add_event( e )

                            dtime = dtime + timedelta( hours=1 )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, localize( dtime ),
                                                  self.default_room )
                            self.add_event( e )

                        else:
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, p.last_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, localize( dtime ), room )
                            self.add_event( e )

                    else:
                        p.recover( )

            self.events.sort( )

        def parse_events_2014(self):

            # START <day> ( <event> <time> { PLUS | MINUS <end time> }? ( AT <room> )? )*
            # START <day> ( <day>* <event> <time> PLUS? )? <day>* <event> <time> PLUS?
            # [ AT <room> ] }+
            # If there are multiple days, then all of the events happen on all of the days (except demos)
            # If there are multiple times, then each represents a separate event on that day

            LOG.debug( 'Parsing events ...' )

            awards_are_events = self.code in ['UPF', 'VIP', 'WWR']
            two_weekends = self.code in ['AFK', 'AOR', 'AUC', 'BWD', 'GBG', 'PZB', 'TRC', 'SQL', 'WAT', 'WSM']

            self.events = []

            weekend_offset = 0
            weekend_start = 5 if two_weekends else 1

            p = Parser( self.event_tokens )

            if self.code in ['WAW']:
                pass

            while p.have_start( ):
                p.match( Token.START )

                days = []
                partial = []

                while p.count and not p.have_start( ):
                    if p.match_day( ):
                        # add day to day queue
                        day = p.last_day
                        days.append( day )

                        weekend_offset = 7 if day.value > weekend_start else weekend_offset

                        day_of_week = day.value + weekend_offset if day.value < 2 else day.value
                        midnight = self.first_day + timedelta( days=day_of_week )

                    elif p.match_multiple_event_times( ):
                        for etime in p.time_list:
                            # handle events immediately
                            dtime = midnight + etime
                            room = self.find_room( p.last_actual, dtime, p.last_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, localize( dtime ), room )
                            self.add_event( e )

                    elif p.match_single_event_time( awards_are_events=awards_are_events ):
                        if p.last_name == 'Demo':
                            # handle demos immediately
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, p.last_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, localize( dtime ), room )
                            self.add_event( e )

                        elif self.code == 'PDT' and p.last_name.endswith( 'FC' ):
                            dtime = midnight + p.last_start
                            room = self.find_room( p.last_actual, dtime, self.draft_room )
                            e = WbcPreview.Event( self.code, self.name, p.last_name + u' Draft', localize( dtime ),
                                                  room )
                            self.add_event( e )

                            dtime = dtime + timedelta( hours=1 )
                            e = WbcPreview.Event( self.code, self.name, p.last_name, localize( dtime ),
                                                  self.default_room )
                            self.add_event( e )

                        else:
                            # add event to event queue
                            e = WbcPreview.Event( self.code, self.name, p.last_name, p.last_start, p.last_room )
                            e.actual = p.last_actual
                            partial.append( e )
                    else:
                        p.recover( )

                for day in days:
                    day_of_week = day.value + weekend_offset if day.value < 2 else day.value
                    midnight = self.first_day + timedelta( days=day_of_week )

                    for pevent in partial:
                        # add event to actual event list
                        dtime = midnight + pevent.time
                        room = self.find_room( pevent.actual, dtime, pevent.location )
                        e = WbcPreview.Event( self.code, self.name, pevent.type, localize( dtime ), room )
                        self.add_event( e )

            self.events.sort( )

        def add_event(self, event):
            if event.code in self.tracking:
                LOG.info( event )
            self.events.append( event )

        def find_room(self, etype, etime, eroom):
            room = self.shift_room if self.shift_room and etime >= self.shift_time else self.default_room
            room = self.event_map[etype] if self.event_map.has_key( etype ) else room
            room = eroom if eroom else room
            room = '-none-' if room is None else room

            if not room:
                pass

            return room

    def __init__(self, metadata):
        self.meta = metadata

        self.valid = False
        self.events = {}

        Token.initialize( )

        LOG.info( "Loading event previews..." )

        for code, url in self.meta.url.items( ):
            # LOG.debug( "Loading event preview for %s: %s", code, url )
            page = parse_url( url )
            if page:
                t = WbcPreview.Tourney( self.meta, code, self.meta.names[code], page )
                self.events[code] = t.events
            else:
                LOG.error( 'Unable to load event preview for %s from %s', code, url )

        self.valid = True
