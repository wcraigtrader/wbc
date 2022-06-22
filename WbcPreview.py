# ----- Copyright (c) 2010-2018 by W. Craig Trader ---------------------------------
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

import logging
import re
from datetime import timedelta, datetime

from bs4 import Tag, NavigableString, Comment

from WbcUtility import parse_url, localize

if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)
    logging.getLogger('requests').setLevel(logging.WARN)

LOG = logging.getLogger('WbcPreview')


# ----- Token -----------------------------------------------------------------

class Token(object):
    """Simple data object for breaking descriptions into parseable tokens"""

    INITIALIZED = False

    CONTINUOUS = PATTERN = DASH = SLASH = PLUS = AT = AND = AM = PM = None

    LOOKUP = {}
    DAYS = {}
    ICONS = {}

    def __init__(self, t, l=None, v=None):
        self.type = t
        self.label = l
        self.value = v

    def __str__(self):
        return str(self.label) if self.label else self.type

    def __repr__(self):
        return self.__str__()

    def __eq__(self, other):
        return self.type == other.type and self.label == other.label and self.value == other.value

    def __ne__(self, other):
        return not self.__eq__(other)

    @classmethod
    def initialize(cls):
        if cls.INITIALIZED:
            return

        cls.INITIALIZED = True

        cls.AM = cls.LOOKUP['AM'] = Token('Meridian', 'AM')
        cls.PM = cls.LOOKUP['PM'] = Token('Meridian', 'PM')

        cls.add_day('SAT1', 0, 'First Saturday')
        cls.add_day('SUN1', 1, 'First Sunday')
        cls.add_day('MON', 2, 'Monday')
        cls.add_day('TUE', 3, 'Tuesday')
        cls.add_day('WED', 4, 'Wednesday')
        cls.add_day('THU', 5, 'Thursday')
        cls.add_day('FRI', 6, 'Friday', 'Fr')
        cls.add_day('SAT2', 7, 'Saturday')
        cls.add_day('SUN2', 8, 'Sunday')

        cls.add_event('Demo', 'Demonstration', 'Demonstrations')
        cls.add_event('H', 'Heat')
        cls.add_event('R', 'Round', 'ound')
        cls.add_event('QF', 'Quarterfinal', 'Quarterfinals')
        cls.add_event('SF', 'Semifinal', 'Semifinals')
        cls.add_event('F', 'Final', 'Finals')
        cls.add_event('Junior')
        cls.add_event('Mulligan', 'mulligan', 'Mulligan Round')
        cls.add_event('After Action', 'After Action Briefing', 'After Action Meeting')
        cls.add_event('Seminar', 'Optional Seminar')
        cls.add_event('Draft', 'DRAFT')
        cls.add_event('Fleet Action')
        cls.add_event('MegaCiv MP Game')

        cls.add_qualifier('PC', 'Grognard PC')
        cls.add_qualifier('AFC')
        cls.add_qualifier('NFC')
        cls.add_qualifier('Super Bowl')

        cls.CONTINUOUS = Token('Symbol', '...')
        cls.LOOKUP['to completion'] = cls.CONTINUOUS
        cls.LOOKUP['till completion'] = cls.CONTINUOUS
        cls.LOOKUP['until completion'] = cls.CONTINUOUS
        cls.LOOKUP['until conclusion'] = cls.CONTINUOUS
        cls.LOOKUP['to conclusion'] = cls.CONTINUOUS

        cls.SHIFT = Token('Symbol', '>')
        cls.LOOKUP['moves to'] = cls.SHIFT
        cls.LOOKUP['moving to'] = cls.SHIFT
        cls.LOOKUP['shifts to'] = cls.SHIFT
        cls.LOOKUP['switches to'] = cls.SHIFT
        cls.LOOKUP['switching to'] = cls.SHIFT
        cls.LOOKUP['after drafts in'] = cls.SHIFT

        cls.LOOKUP['HMWG'] = Token('Format', 'HMWG')

        cls.initialize_7springs_rooms()

        cls.AT = Token('Symbol', '@')
        cls.AND = cls.LOOKUP['and'] = Token('Symbol', '&')
        cls.PLUS = Token('Symbol', '+')
        cls.DASH = Token('Symbol', '-')
        cls.SLASH = Token('Symbol', '/')
        cls.START = Token('Symbol', '|')

        cls.PATTERN = '|'.join(sorted(cls.LOOKUP.keys(), reverse=True))

        cls.LOOKUP['&'] = cls.AND
        cls.LOOKUP['@'] = cls.AT
        cls.LOOKUP['+'] = cls.PLUS
        cls.LOOKUP['-'] = cls.DASH
        cls.LOOKUP['/'] = cls.SLASH
        cls.LOOKUP['|'] = cls.START

        cls.PATTERN += '|[@&+-/]'

    @classmethod
    def add_event(cls, primary, *aliases):
        room = Token('Event', primary)
        cls.LOOKUP[primary] = room
        for alias in aliases:
            cls.LOOKUP[alias] = room

    @classmethod
    def add_qualifier(cls, primary, *aliases):
        room = Token('Qualifier', primary)
        cls.LOOKUP[primary] = room
        for alias in aliases:
            cls.LOOKUP[alias] = room

    @classmethod
    def add_day(cls, name, value, *aliases):
        day = Token('Day', name, value)
        cls.DAYS[name] = day
        cls.LOOKUP[name] = day
        for alias in aliases:
            cls.LOOKUP[alias] = day

    @classmethod
    def add_room(cls, primary, *aliases):
        room = Token('Room', primary)
        cls.LOOKUP[primary] = room
        for alias in aliases:
            cls.LOOKUP[alias] = room

    @classmethod
    def initialize_7springs_rooms(cls):
        cls.add_room('Alpine')
        cls.add_room('Ballroom', 'Ballroom B')
        cls.add_room('Ballroom Stage')
        cls.add_room('Bavarian Lounge')
        cls.add_room('Chestnut')
        cls.add_room('Dogwood', 'Dogwood Forum')
        cls.add_room('Evergreen')
        cls.add_room('Evergreen & Chestnut')
        cls.add_room('Exhibit Annex T1', 'Exhibit Annex Table 1', 'Exhibit annex Table 1')
        cls.add_room('Exhibit Annex T2', 'Exhibit Annex Table 2', 'Exhibit annex Table 2')
        cls.add_room('Exhibit Annex T3', 'Exhibit Annex Table 3', 'Exhibit annex Table 3')
        cls.add_room('Exhibit Annex T4', 'Exhibit Annex Table 4', 'Exhibit annex Table 4')
        cls.add_room('Exhibit Annex T5', 'Exhibit Annex Table 5', 'Exhibit annex Table 5')
        cls.add_room('Exhibit Annex T6', 'Exhibit Annex Table 6', 'Exhibit annex Table 6')
        cls.add_room('Exhibit Annex T7', 'Exhibit Annex Table 7', 'Exhibit annex Table 7')
        cls.add_room('Exhibit Annex T8', 'Exhibit Annex Table 8', 'Exhibit annex Table 8')
        cls.add_room('Exhibit Annex T9', 'Exhibit Annex Table 9', 'Exhibit annex Table 9')
        cls.add_room('Exhibit Hall')
        cls.add_room('Winterberry', 'Festival Hall', 'Festival')
        cls.add_room('First Tracks Center', 'Ski Lodge First Tracks Center')
        cls.add_room('First Tracks Poolside', 'Ski Lodge First Tracks Poolside')
        cls.add_room('First Tracks Slopeside', 'Ski Lodge First Tracks Slopeside')
        cls.add_room('Foggy Brews')
        cls.add_room('Foggy Goggle Center', 'Ski Lodge Foggy Goggle Center')
        cls.add_room('Foggy Goggle Front', 'Ski Lodge Foggy Goggle Front')
        cls.add_room('Foggy Goggle Rear', 'Ski Lodge Foggy Goggle Rear')
        cls.add_room('Fox Den')
        cls.add_room('Hemlock')
        cls.add_room('Laurel')
        cls.add_room('Maple Room', 'Maple', 'Ski Lodge Maple Room')
        cls.add_room('Rathskeller')
        cls.add_room('Seasons')
        cls.add_room('Seasons 1')
        cls.add_room('Seasons 1-2')
        cls.add_room('Seasons 1-3')
        cls.add_room('Seasons 1-4')
        cls.add_room('Seasons 1-5')
        cls.add_room('Seasons 2')
        cls.add_room('Seasons 2-3')
        cls.add_room('Seasons 2-4')
        cls.add_room('Seasons 2-5')
        cls.add_room('Seasons 3')
        cls.add_room('Seasons 3-4')
        cls.add_room('Seasons 3-5')
        cls.add_room('Seasons 4')
        cls.add_room('Seasons 4-5')
        cls.add_room('Seasons 5')
        cls.add_room('Snowflake Forum', 'Snowflake')
        cls.add_room('Stag Pass')
        cls.add_room('Sunburst Forum')
        cls.add_room('Timberstone')
        cls.add_room('Wintergreen')

        # Misspelled Room Names
        cls.add_room('Ski Lodge Fast Tracks Center')
        cls.add_room('Ski Lodge Foggie Goggle Front')
        cls.add_room('Fast Tracks Slopeside')

    @classmethod
    def initialize_host_rooms(cls):
        cls.LOOKUP['Ballroom A'] = Token('Room', 'Ballroom A')
        cls.LOOKUP['Ballroom B'] = Token('Room', 'Ballroom B')
        cls.LOOKUP['Ballroom AB'] = Token('Room', 'Ballroom AB')
        cls.LOOKUP['Ballroom'] = cls.LOOKUP['Ballroom AB']

        cls.LOOKUP['Conestoga 1'] = Token('Room', 'Conestoga 1')
        cls.LOOKUP['Conestoga 2'] = Token('Room', 'Conestoga 2')
        cls.LOOKUP['Conestoga 3'] = Token('Room', 'Conestoga 3')
        cls.LOOKUP['Coonestoga 3'] = cls.LOOKUP['Conestoga 3']

        cls.LOOKUP['Cornwall'] = Token('Room', 'Cornwall')
        cls.LOOKUP['Cromwell'] = cls.LOOKUP['Cornwall']
        cls.LOOKUP['Heritage'] = Token('Room', 'Heritage')
        cls.LOOKUP['Hopewell'] = Token('Room', 'Hopewell')
        cls.LOOKUP['Kinderhook'] = Token('Room', 'Kinderhook')
        cls.LOOKUP['Lampeter'] = Token('Room', 'Lampeter')
        cls.LOOKUP['Laurel Grove'] = Token('Room', 'Laurel Grove')
        cls.LOOKUP['Limerock'] = Token('Room', 'Limerock')
        cls.LOOKUP['Marietta'] = Token('Room', 'Marietta')
        cls.LOOKUP['New Holland'] = Token('Room', 'New Holland')
        cls.LOOKUP['Paradise'] = Token('Room', 'Paradise')
        cls.LOOKUP['Showroom'] = Token('Room', 'Showroom')
        cls.LOOKUP['Strasburg'] = Token('Room', 'Strasburg')
        cls.LOOKUP['Wheatland'] = Token('Room', 'Wheatland')

        cls.LOOKUP['Terrace 1'] = Token('Room', 'Terrace 1')
        cls.LOOKUP['Terrace 2'] = Token('Room', 'Terrace 2')
        cls.LOOKUP['Terrace 3'] = Token('Room', 'Terrace 3')
        cls.LOOKUP['Terrace 4'] = Token('Room', 'Terrace 4')
        cls.LOOKUP['Terrace 5'] = Token('Room', 'Terrace 5')
        cls.LOOKUP['Terrace 6'] = Token('Room', 'Terrace 6')
        cls.LOOKUP['Terrace 7'] = Token('Room', 'Terrace 7')

        cls.LOOKUP['Vista C'] = Token('Room', 'Vista C')
        cls.LOOKUP['Vista D'] = Token('Room', 'Vista D')
        cls.LOOKUP['Vista CD'] = Token('Room', 'Vista CD')
        cls.LOOKUP['Vista'] = cls.LOOKUP['Vista CD']

    @classmethod
    def initialize_icons(cls):
        cls.ICONS = {
            'semi': Token('Award', 'SF'),
            'final': Token('Award', 'F'),
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
            'sty_cont': cls.CONTINUOUS,
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
            'for_mese': Token('Format', 'Heats'),
            'for_se': Token('Format', 'Single Elimination'),
            'for_sem': Token('Format', 'Single Elimination Mulligan'),
            'for_swis': Token('Format', 'Swiss'),
            'for_swel': Token('Format', 'Swiss Elimination'),
            'freeform': Token('Style', 'Freeform'),
            'sty_sche': Token('Style', 'Scheduled'),
            'sty_ctht': Token('Style', 'Continuous Heats'),
            'prize1': Token('Prizes', 1, 1),
            'prize2': Token('Prizes', 2, 2),
            'prize3': Token('Prizes', 3, 3),
            'prize4': Token('Prizes', 4, 4),
            'prize5': Token('Prizes', 5, 5),
            'prize6': Token('Prizes', 6, 6),
            'prztrial': Token('Prizes', 'Trial', 0),
        }

    @staticmethod
    def encode_list(tokenlist):
        return [u' '.join([unicode(y) for y in x]) for x in tokenlist]

    @classmethod
    def tokenize(cls, tag):
        tokens = []
        partial = u''

        for tag in tag.descendants:
            if isinstance(tag, Comment):
                pass  # Always ignore comments
            elif isinstance(tag, NavigableString):
                partial += u' ' + unicode(tag)
            elif isinstance(tag, Tag) and tag.name in ['img']:
                tokens += cls.tokenize_text(partial)
                partial = u''
                tokens += cls.tokenize_icon(tag)
            else:
                pass  # ignore other tags, for now
                # LOG.debug( 'Ignored <%s>', tag.name )

        if partial:
            tokens += cls.tokenize_text(partial)

        return tokens

    @classmethod
    def tokenize_icon(cls, tag):
        cls.initialize()

        tokens = []

        try:
            name = tag['src'].lower()
            name = name.split('/')[-1]
            name = name.split('.')[0]
        except:
            LOG.error("%s didn't have a 'src' attribute", tag)
            return tokens

        if name in cls.ICONS:
            tokens.append(cls.ICONS[name])
        elif name in ['stadium', 'class_a', 'class_b', 'coached']:
            pass
        elif name.startswith('for_'):
            form = name[4:]
            token = Token('Format', form)
            cls.ICONS[name] = token
            tokens.append(token)
            LOG.warn('Automatically added form [%s]', form)
        elif name.startswith('sty_'):
            style = name[4:]
            token = Token('Style', style)
            cls.ICONS[name] = token
            tokens.append(token)
            LOG.warn('Automatically added style [%s]', style)
        else:
            LOG.warn('Ignored icon [%s]', name)

        return tokens

    @classmethod
    def tokenize_text(cls, text):
        cls.initialize()

        data = text

        junk = u''
        tokens = []

        # Cleanup crappy data
        data = data.replace(u'\xa0', u' ')
        data = data.replace(u'\n', u' ')

        data = data.strip()
        data = data.replace(u' ' * 11, u' ').replace(u' ' * 7, u' ').replace(u' ' * 5, u' ')
        data = data.replace(u' ' * 3, u' ').replace(u' ' * 2, u' ')
        data = data.replace(u' ' * 2, u' ').replace(u' ' * 2, u' ')
        data = data.strip()

        hdata = data.encode('unicode_escape')

        while len(data):

            # Ignore commas and semi-colons
            if data[0] in u',;:':
                data = data[1:]
            else:
                # Match Room names, event names, phrases, symbols
                m = re.match(cls.PATTERN, data)
                if m:
                    text = m.group()
                    tokens.append(cls.LOOKUP[text])
                    data = data[len(text):]
                else:
                    # Match numbers
                    m = re.match("\d+", data)
                    if m:
                        text = m.group()
                        n = int(text)
                        tokens.append(Token('Number', text, n))
                        data = data[len(text):]
                    else:
                        junk += data[0]
                        data = data[1:]

            data = data.strip()

        if junk:
            hjunk = junk.encode('unicode_escape')
            LOG.log(logging.NOTSET, 'Skipped [%s] in [%s]', hjunk, hdata)

        return tokens


class Parser(object):
    tokens = []
    last_match = None

    def __init__(self, tokens):
        self.tokens = list(tokens)

    def __str__(self):
        return "%s (%d) %s" % (self.last_match, self.count, self.tokens)

    @property
    def count(self):
        return len(self.tokens)

    def has(self, *prediction):
        lookahead_length = len(prediction)
        if self.count < lookahead_length:
            return False

        pos = 0
        for predicted in prediction:
            lookahead = self.tokens[pos]
            pos += 1

            if isinstance(predicted, Token):
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
            if isinstance(stop, Token):
                if current == stop:
                    return False
            elif current.type == stop:
                return False
        return True

    def next(self):
        return self.tokens.pop(0)

    def match_3_events(self):
        """
        Match 'Event', Token.SLASH, 'Event', Token.SLASH, 'Event'
        Stop [ 'Day' ]
        """
        e1 = self.next()
        self.next()
        e2 = self.next()
        self.next()
        e3 = self.next()
        return [(e1.label,), (e2.label,), (e3.label,)]

    def match_2_events(self):
        """
        Match 'Event', Token.SLASH, 'Event'
        Stop [ 'Day' ]
        """
        e1 = self.next()
        self.next()
        e2 = self.next()
        return [(e1.label,), (e2.label,)]

    def match_event(self):
        """
        Match 'Event'
        Stop [ 'Day' ]
        """
        e1 = self.next()
        return (e1.label,)

    def match_qualified_event(self):
        """
        Match 'Qualifier' 'Event'
        Stop [ 'Day' ]
        """
        q = self.next()
        e = self.next()
        return e.label, q.label

    def match_multiple_heats(self):
        """
        Match 'Event', 'Number', Token.DASH, 'Number'
        Stop [ 'Day' ]
        """
        e1 = self.next()
        first = self.next()
        self.next()
        last = self.next()

        results = []
        for n in range(first.value, last.value + 1):
            results.append((e1.label, n))

        return results

    def match_heat(self):
        """
        Match 'Event', 'Number', 'Qualifier'
        Match 'Event', 'Number'
        Stop [ 'Day' ]
        """
        e1 = self.next()
        n = self.next()
        if self.has('Qualifier'):
            q = self.next()
            return e1.label, n.value, q.label

        return e1.label, n.value

    def match_round(self):
        """
        Match 'Event', 'Number', Token.SLASH, 'Number'
        Stop [ 'Day' ]
        """
        e1 = self.next()
        n = self.next()
        self.next()
        m = self.next()

        return e1.label, n.value, m.value

    def match_date_times(self):
        """Recognized date/time formats:
    
        <day> <time> @
        <day> @ <time>
        <day> @ <time> - <time>
        <day> <time> & <day> <time> @
        <day> <time> <time> <time> & <time> @
        <day> <time> & <day> <time> & <time> & <day> <time> & <day> <time> @
        <day> @ <time> <time> <time> <time>
    
        Ignores [ &, @ ]
        Stops on [ -?,  <room> ]
    
        Returns a list of ( Day, Time ) tuples
        """

        results = []
        while self.has('Day'):
            day = self.next()
            while self.is_not(Token.DASH, 'Room', 'Day'):
                if self.has('Number', Token.DASH, 'Number'):
                    time = self.next()
                    self.next()
                    self.next()
                    results.append(timedelta(days=day.value, hours=time.value))
                if self.has('Number'):
                    offset = 0
                    time = self.next()
                    if self.has(Token.AM):
                        self.next()
                    elif self.has(Token.PM):
                        self.next()
                        offset = 12
                    elif self.has(Token.PLUS):
                        self.next()  # FIXME: Do something with + => continuous???
                    results.append(timedelta(days=day.value, hours=time.value + offset))
                elif self.has(Token.AND) or self.has(Token.AT):
                    self.next()

        return results

    def match_room(self):
        if self.count and self.tokens[0] == Token.DASH:
            self.next()

        if self.count and self.tokens[0].type == 'Room':
            return self.next()

        return None


# ----- WBC Preview Schedule -------------------------------------------------

class WbcPreview(object):
    """This class is used to parse schedule data from the annual preview pages"""

    tracking = [
    ]

    notes = {}

    class Event(object):
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
            return cmp(self.time, other.time)

        def __str__(self):
            return '%s %s %s in %s at %s' % (self.code, self.name, self.type, self.location, self.time)

    class Tourney(object):

        def __init__(self, metax, code, name, page):

            self.meta = metax
            self.code = code
            self.name = name

            self.notes = []
            self.events = []
            self.event_tokens = []
            self.heats = set()
            self.rounds = set()
            self.max = 0

            td = page.table.table.table.findAll('tr')[2].td
            paras = list(td.findAll('p'))

            if len(paras):
                self.tokenize_events(paras)
                self.parse_events()
                self.check_consistency()

            else:
                LOG.error('%s: Did not find schedule data', code)

        def tokenize_events(self, paras):
            for para in paras:
                text = para.text.strip()
                for line in text.split('\n'):
                    stripped = line.strip()
                    if stripped:
                        self.event_tokens.append(Token.tokenize_text(stripped))

            if self.code in WbcPreview.tracking:
                for section in Token.encode_list(self.event_tokens):
                    LOG.debug("%s: %s", self.code, section)

        def parse_events(self):
            """Recognized event formats:

            QF / SF / F <day> @ <time> -? <room>
            QF / SF <day> @ <time> -? <room>
            SF / F <day> @ <time> -? <room>
            <event> # - # <day> @ <time> <time> <time> <time> -? <room>
            <event> # / # <day> @ <time> -? <room>
            <event> # <qualifier> <day> @ <time> -? <room>
            <event> # <day> @ <time> -? <room>
            <event> <day> @ <time> -? <room>

            Demo <day> <time> @ <room>
            Demo <day> @ <time> -? <room>
            Demo <day> <time> & <day> <time> @ <room>
            Demo <day> <time> <time> <time> & <time> @ <room>
            Demo <day> <time> & <day> <time> & <time> & <day> <time> & <day> <time> @ <room>
            """

            row_number = 0
            for row in self.event_tokens:
                LOG.debug("Tourney %s Parsing row %d: %s", self.code, row_number, row)
                row_number += 1

                p = Parser(row)
                if p.has('Event', Token.SLASH, 'Event', Token.SLASH, 'Event', 'Day'):
                    elist = p.match_3_events()
                    times = p.match_date_times()
                    room = p.match_room()
                    self.add_events([e[0] for e in elist], times, room)
                    self.max = max(self.max, 4)
                    self.rounds.add(self.max - 2)
                    self.rounds.add(self.max - 1)
                    self.rounds.add(self.max)

                elif p.has('Event', Token.SLASH, 'Event', 'Day'):
                    elist = p.match_2_events()
                    times = p.match_date_times()
                    room = p.match_room()
                    self.add_events([e[0] for e in elist], times, room)
                    if elist[0][0] == 'QF':
                        self.max = max(self.max, 4)
                        self.rounds.add(self.max - 2)
                        self.rounds.add(self.max - 1)
                    elif elist[0][0] == 'SF':
                        self.max = max(self.max, 3)
                        self.rounds.add(self.max - 1)
                        self.rounds.add(self.max)

                elif p.has('Event', 'Number', Token.SLASH, 'Number', 'Event', 'Number', 'Day'):
                    round, n, m = p.match_round()
                    heat, h = p.match_heat()
                    times = p.match_date_times()
                    room = p.match_room()

                    if round == 'R' and heat == 'H' and n == 1:
                        ename = '%s%d' % (heat, h)
                        if len(times) == 0:
                            self.notes.append('Preview missing start time for %s in %s' % (ename, room))
                        elif len(times) > 1:
                            self.notes.append('Preview has extra start times for %s in %s' % (ename, room))
                        else:
                            self.add_event(ename, times[0], room)
                        self.max = max(self.max, 2)
                        self.rounds.add(1)
                        if n in self.heats:
                            self.notes.append('Preview has duplicate entry for %s' % ename)
                        else:
                            self.heats.add(h)

                elif p.has('Event', 'Number', Token.DASH, 'Number', 'Day'):
                    elist = p.match_multiple_heats()
                    times = p.match_date_times()
                    room = p.match_room()
                    self.add_events(["%s%d" % event for event in elist], times, room)
                    self.max = max(self.max, 2)
                    self.rounds.add(1)

                elif p.has('Event', 'Number', Token.SLASH, 'Number', 'Day'):
                    name, n, m = p.match_round()
                    times = p.match_date_times()
                    room = p.match_room()

                    if name == 'R':
                        ename = '%s%d/%d' % (name, n, m)
                        self.max = max(self.max, m)
                        if n in self.rounds:
                            self.notes.append('Preview has duplicate entry for %s' % ename)
                        else:
                            self.rounds.add(n)
                    elif name == 'H':
                        ename = '%s%d/%d' % (name, n, m)
                        self.max = max(self.max, 2)
                        self.rounds.add(1)
                        if n in self.heats:
                            self.notes.append('Preview has duplicate entry for %s' % ename)
                        else:
                            self.heats.add(n)
                    else:
                        ename = '%s %d/%d' % (name, n, m)

                    self.add_event(ename, times[0], room)

                elif p.has('Event', 'Number', 'Qualifier', 'Day'):
                    name, n, qualifier = p.match_heat()
                    times = p.match_date_times()
                    room = p.match_room()
                    ename = '%s %d %s' % (name, n, qualifier)
                    self.add_event(ename, times[0], room)
                    if name == 'H':
                        self.max = max(self.max, 2)
                        self.rounds.add(1)
                        if n in self.heats:
                            self.notes.append('Preview has duplicate entry for %s' % ename)
                        else:
                            self.heats.add(n)

                elif p.has('Event', 'Number', 'Day'):
                    name, n = p.match_heat()
                    times = p.match_date_times()
                    room = p.match_room()
                    ename = '%s%d' % (name, n)
                    if len(times) == 0:
                        self.notes.append('Preview missing start time for %s in %s' % (ename, room))
                    elif len(times) > 1:
                        self.notes.append('Preview has extra start times for %s in %s' % (ename, room))
                    else:
                        self.add_event(ename, times[0], room)

                    if name == 'H':
                        self.max = max(self.max, 2)
                        self.rounds.add(1)
                        if n in self.heats:
                            self.notes.append('Preview has duplicate entry for %s' % ename)
                        else:
                            self.heats.add(n)

                elif p.has('Qualifier', 'Event', 'Day'):
                    event = p.match_qualified_event()
                    times = p.match_date_times()
                    room = p.match_room()
                    ename = ' '.join(event)
                    if len(times) == 0:
                        self.notes.append('Preview missing start time for %s in %s' % (ename, room))
                    elif len(times) > 1:
                        self.notes.append('Preview has extra start times for %s in %s' % (ename, room))
                    else:
                        self.add_event(ename, times[0], room)


                elif p.has('Event', 'Day'):
                    event = p.match_event()
                    times = p.match_date_times()
                    room = p.match_room()

                    count = len(times)
                    name = event[0]
                    if count > 1:
                        names = ["%s %d/%d" % (name, i, count) for i in range(1, count + 1)]
                        self.add_events(names, times, room)
                    else:
                        self.add_event(name, times[0], room)

                    if name == 'QF':
                        self.max = max(self.max, 4)
                        self.rounds.add(self.max - 2)
                    elif name == 'SF':
                        self.max = max(self.max, 3)
                        self.rounds.add(self.max - 1)
                    elif name == 'F':
                        self.max = max(self.max, 2)
                        self.rounds.add(self.max)

                elif p.has(Token.DASH):
                    pass  # Do nothing

                else:
                    LOG.error('%s: Could not match %s', self.code, row)

            self.events.sort()

        def add_event(self, name, time, room):
            event_time = localize(self.meta.first_day + time)
            if room is None:
                self.notes.append('Preview missing room for %s at %s' % (name, event_time))
            else:
                event = WbcPreview.Event(self.code, self.name, name, event_time, room.label)
                self.events.append(event)

        def add_events(self, names, times, room):
            for i in range(len(times), len(names)):
                self.notes.append('Preview missing start time for %s in %s' % (names[i], room.label))

            for name, time in zip(names, times):
                self.add_event(name, time, room)

        def check_consistency(self):
            if self.max:
                missing = list(set(range(1, self.max + 1)) - self.rounds)
                missing.sort()
                for r in missing:
                    self.notes.append('Preview missing start time for R%d/%d' % (r, self.max))

    def __init__(self, metadata):
        self.meta = metadata
        self.tracking.extend(metadata.tracking)

        self.valid = False
        self.events = {}

        Token.initialize()

        LOG.info("Loading event previews...")
        LOG.debug("Assuming first day is %s", self.meta.first_day)

        for code, url in sorted( self.meta.url.items() ):
            LOG.debug("Loading event preview for [%s]: %s", code, url)
            if len(self.tracking) and code not in self.tracking:
                continue

            if code not in self.notes:
                self.notes[code] = []

            page = parse_url(url)
            if page:
                t = WbcPreview.Tourney(self.meta, code, self.meta.names[code], page)
                self.events[code] = t.events
                self.notes[code].extend(t.notes)
            else:
                message = 'Unable to load event preview for %s from %s' % (code, url)
                self.notes[code].append(message)
                LOG.error(message)

        self.valid = True


# ----- Testing --------------------------------------------------------------

if __name__ == '__main__':
    from WbcMetadata import WbcMetadata

    meta = WbcMetadata()
    meta.first_day = datetime(2017, 7, 22, 0, 0, 0)
    preview = WbcPreview(meta)
