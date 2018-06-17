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
import unicodedata
from collections import OrderedDict
from datetime import datetime, time, timedelta

import xlrd

from WbcSpreadsheet import WbcRow, WbcSchedule
from WbcUtility import round_up_timedelta

if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)
    logging.getLogger('requests').setLevel(logging.WARN)

LOG = logging.getLogger('WbcOldSpreadsheet')


# ----- WBC Old Excel Row (read from Excel spreadsheet) -----------------------

class WbcOldRow(WbcRow):
    """This subclass of WbcRow is used to parse Excel-formatted schedule data"""

    KEYS = ['Date', 'Time', 'Event', 'Prize', 'Class', 'Format', 'Duration', 'C', 'GM', 'Location']
    FIELDS = ['Date', 'Time', 'Event', 'Prize', 'Class', 'Format', 'Duration', 'Continuous', 'GM', 'Location']
    GENERATED = ['Code']

    def __init__(self, *args):
        self.keymap = OrderedDict(zip(self.KEYS, self.FIELDS))
        WbcRow.__init__(self, *args)

    def readrow(self, *args):
        """Custom implementation of readrow to handle XLS-formatted rows"""

        labels = args[0]
        row = args[1]
        datemode = args[2]

        for i in range(len(labels)):
            key = labels[i]

            val = row[i]
            if not key:
                continue
            elif val.ctype == xlrd.XL_CELL_EMPTY:
                val = None
            elif val.ctype == xlrd.XL_CELL_TEXT:
                val = unicodedata.normalize('NFKD', val.value).encode('ascii', 'ignore').strip()
            elif val.ctype == xlrd.XL_CELL_NUMBER:
                val = str(float(val.value))
            elif val.ctype == xlrd.XL_CELL_DATE:
                val = xlrd.xldate_as_tuple(val.value, datemode)
                if val[0]:
                    val = datetime(*val)  # pylint: disable=W0142
                else:
                    val = time(val[3], val[4], val[5])
            else:
                raise ValueError("Unhandled Excel cell type (%s) for %s" % (val.ctype, key))

            self.__setattr__(key, val)

        self.name = self.event.strip()

    def initialize(self):
        # Check for errors that will throw exceptions later
        if self.name.endswith('Final'):  # Replace trailing 'Final' with 'F'
            self.name = self.name[:-4]

        if self.name.endswith(' MWR'):  # MWR = Mulligan Winner Round, but that's not important right now.
            self.name = self.name[:-4]

        # parse the data to generate useful fields
        self.cleanlocation()
        self.checkrounds()
        self.checktypes(self.schedule.TYPES)
        self.checkrounds()
        self.checktypes(self.schedule.JUNIOR)
        self.checktimes()
        self.checkduration()
        self.checkcodes()

    def checktimes(self):
        """Custom implementation of checktimes to handle XLS-formatted date/times"""

        d = self.date

        if self.time.__class__ is time:
            t = self.time
        else:
            try:
                t = float(self.time)
                if t > 23:
                    t = time(23, 59)
                    self.length = timedelta(minutes=1)
                else:
                    m = t * 60
                    h = int(m / 60)
                    m = int(m % 60)
                    t = time(h, m)
            except:
                try:
                    t = datetime.strptime(self.time, "%H:%M")
                except:
                    try:
                        t = datetime.strptime(self.time, "%I:%M:%S %p")
                    except:
                        raise ValueError('Unable to parse (%s) as a time' % self.time)

        self.datetime = d.replace(hour=t.hour, minute=t.minute)

    def checkrounds(self):
        """Check the current state of the event name to see if it describes a Heat or Round number"""

        match = re.search(r'([DSHR]?)(\d+)[-/](\d+)$', self.name)
        if match:
            (t, n, m) = match.groups()
            text = match.group(0)
            text = text.replace('-', '/')
            if t == "R":
                self.start = int(n)
                self.rounds = int(m)
                self.rtype = text.strip()
                self.name = self.name[:-len(text)].strip()
            elif t == "H" or t == '':
                self.type = self.type + ' ' + text
                self.type = self.type.strip()
                self.name = self.name[:-len(text)].strip()
            elif t == "D":
                dtext = text.replace('D', '')
                self.type = self.type + ' Demo ' + dtext
                self.type = self.type.strip()
                self.name = self.name[:-len(text)].strip()
            elif t == "P":
                dtext = text.replace('P', '')
                self.type = self.type + ' Preview ' + dtext
                self.type = self.type.strip()
                self.name = self.name[:-len(text)].strip()

    def checktypes(self, types):
        """
        Check the current state of the event name and strip off ( and flag ) any of the listed
        event type codes
        """

        for event_type in types:
            if self.name.endswith(event_type):
                self.type = event_type + ' ' + self.type
                self.name = self.name[:-len(event_type)].strip()
                if event_type == 'FF':
                    self.freeformat = True
                elif event_type == 'PC':
                    self.grognard = True
                elif event_type in self.schedule.JUNIOR:
                    self.junior = True

        if self.name.startswith('JR '):
            self.type = 'JR ' + self.type
            self.name = self.name[3:].strip()
            self.junior = True

        self.type = self.type.strip()

    def checkduration(self):
        """
        Given the current event state, set the continuous event flag,
        and calculate the correct event length.
        """

        if 'continuous' in self.__dict__:
            self.continuous = (self.continuous in ('C', 'Y'))
        else:
            self.continuous = False

        if not self.duration or self.duration == '-':
            return

        if self.duration.endswith("q"):
            self.continuous = True
            self.duration = self.duration[:-1]

        m = re.match("(\d+)\[(\d+)\]", self.duration)
        if m:
            self.duration = m.group(1)

        try:
            l = timedelta(minutes=60 * float(self.duration))
            self.length = round_up_timedelta(l)
        except:
            LOG.error("Invalid duration (%s) on %s", self.duration, self)
            self.length = timedelta(minutes=0)

    def checkcodes(self):
        """
        Check the current state of the event name and identify the actual event that matches
        the abbreviated name that's present.
        """

        self.code = None

        # First check for Junior events
        if self.junior:
            self.code = 'junior'

        # Check for tournament codes
        elif self.name in self.meta.codes:
            self.code = self.meta.codes[self.name]
            self.name = self.meta.names[self.code]

            # If this event has rounds, save them for later use
            if self.rounds:
                self.schedule.rounds[self.code] = self.rounds

        else:
            # Check for non-tournament groupings
            for o in self.meta.others:
                if ((o['format'] and o['format'] == self.format) or
                        (o['name'] and o['name'] == self.name)):
                    self.code = o['code']
                    LOG.debug("Other (%s) %s | %s", o['code'], o['name'], o['format'])
                    return

    def cleanlocation(self):
        """Clean up typical typos in the location name"""
        if self.location:
            self.location = self.location.strip()
            self.location = self.location.replace('  ', ' ')
            self.location = self.location.replace('Marieta', 'Marietta')


class WbcOldSchedule(WbcSchedule):
    """
    Up through 2017, the spreadsheet was an Excel file that had a single header line,
    that needed crazy heuristics to (more or less) generate a real schedule.
    """

    # Recognized event flags
    FLAVOR = ['AFC', 'NFC', 'FF', 'Circus', 'DDerby', 'Draft', 'Playoffs', 'FF']
    JUNIOR = ['Jr', 'Jr.', 'Junior', 'JR']
    TEEN = ['Teen']
    MENTOR = ['Mentoring']
    MULTIPLE = ['QF/SF/F', 'QF/SF', 'SF/F']
    SINGLE = ['QF', 'SF', 'F']
    STYLE = ['After Action Debriefing', 'After Action Meeting', 'After Action', 'Aftermath', 'Awards', 'Demo',
             'Mulligan', 'Preview'] + MULTIPLE + SINGLE

    TYPES = ['PC'] + FLAVOR + JUNIOR + TEEN + MENTOR + STYLE

    def __init__(self, *args):
        WbcSchedule.__init__(self, *args)

    def load_events(self):
        """
        Read an Excel spreadsheet and generate WBC events for each row.
        """

        if not self.filename:
            self.filename = "schedule%s.xlsx" % self.meta.year

        LOG.debug('Reading Excel spreadsheet from %s', self.filename)

        book = xlrd.open_workbook(self.filename)
        sheet = book.sheet_by_index(0)

        header = []
        for i in range(sheet.ncols):
            key = None
            try:
                key = sheet.cell_value(0, i)
                if key:
                    key = unicodedata.normalize('NFKD', key).encode('ascii', 'ignore').lower()
                header.append(key)

            except Exception:
                raise ValueError('Unable to parse Column Header %d (%s)' % (i, key))

        for i in range(1, sheet.nrows):
            if self.meta.verbose:
                LOG.debug('Reading row %d' % (i + 1))
            try:
                event_row = WbcOldRow(self, i + 1, header, sheet.row(i), book.datemode)
                self.categorize_row(event_row)
            except Exception as e:
                LOG.error('Skipped row %d: %s', i + 1, e.message)

        for event_list in self.events.values():
            event_list.sort(lambda x, y: cmp(x.datetime, y.datetime))
            for event in event_list:
                self.process_event(event)

    def process_event(self, entry):
        """
        For a spreadsheet entry, generate calendar events as follows:

        If the entry is WAW,
            treat it like an all-week free-format event.
        If the entry is free-format, and a grognard,
            it's an all-week event, but use the grognard duration from the event codes.
        If the entry is free-format, and a Swiss Elimination, and it has rounds,
            it's an all-week event, but use the duration from the event codes.
        If the entry has rounds,
            add calendar events for each round.
        If the entry is marked as continuous, and it's marked 'HMSE',
           there's no clue as to how many actual heats there are, so just add one event.
        If the entry is marked as continuous,
            add as many events as there are types coded.
        Otherwise,
           add a single event for the entry.
        """

        calendar = self.get_or_create_event_calendar(entry.code)
        eventmeta = self.meta.eventmeta.get(entry.code, None)
        grognard = eventmeta and eventmeta.grognard

        # This test is for debugging purposes, and is only good for an entry that was sucessfully coded
        if entry.code in self.tracking:
            pass

        if entry.code == 'WAW' and entry.format != 'Meeting':
            self.process_all_week_entry(calendar, entry)
        elif entry.freeformat and entry.grognard:
            self.process_freeformat_grognard_entry(calendar, entry)
        elif grognard and entry.format == 'SwEl':
            self.process_freeformat_grognard_entry(calendar, entry)
        elif entry.freeformat and entry.format == 'SwEl' and entry.rounds:
            self.process_freeformat_swel_entry(calendar, entry)
        elif entry.rounds:
            self.process_entry_with_rounds(calendar, entry)
        elif entry.continuous and entry.format == 'HMSE':
            self.process_normal_entry(calendar, entry)
        elif entry.continuous:
            self.process_continuous_event(calendar, entry)
        else:
            self.process_normal_entry(calendar, entry)

    def process_normal_entry(self, calendar, entry):
        """
        Process a spreadsheet entry that maps to a single event.
        """
        name = entry.name + ' ' + entry.type
        alternative = self.alternate_round_name(entry)
        if alternative:
            self.add_event(calendar, entry, name=name, altname=alternative)
        else:
            self.add_event(calendar, entry, name=name, replace=False)

    def process_continuous_event(self, calendar, entry):
        """
        Process multiple back-to-back events that are not rounds, per se.
        """
        start = entry.datetime
        for event_type in entry.type.split('/'):
            name = entry.name + ' ' + event_type
            alternative = self.alternate_round_name(entry, event_type)
            self.add_event(calendar, entry, start=start, name=name, altname=alternative)
            start = self.calculate_next_start_time(entry, start)

    def process_entry_with_rounds(self, calendar, entry):
        """
        Process multiple back-to-back rounds
        """
        start = entry.datetime
        name = entry.name + ' ' + entry.type
        name = name.strip()

        rounds = range(int(entry.start), int(entry.rounds) + 1)
        for r in rounds:
            label = "%s R%s/%s" % (name, r, entry.rounds)
            self.add_event(calendar, entry, start=start, name=label)
            start = self.calculate_next_start_time(entry, start)

    def calculate_next_start_time(self, entry, start):
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
        midnight = datetime.fromordinal(start.toordinal() + 1)

        # Calculate 9am tomorrow
        tomorrow = datetime(midnight.year, midnight.month, midnight.day, 9, 0, 0)

        # Nominal start time and end time for the next event
        next_start = start + entry.length
        next_end = next_start + entry.length

        late_part = 0.0
        if next_end > midnight:
            late_part = (next_end - midnight).total_seconds() / entry.length.total_seconds()

        # Lookup the override code for this event
        eventmeta = self.meta.eventmeta.get(entry.code, None)
        playlate = eventmeta.playlate if eventmeta else None

        if playlate and late_part:
            LOG.debug("Play late: %s: %4s, Start: %s, End: %s, Partial: %5.2f, %s", entry.code, playlate, next_start,
                      next_end, late_part, late_part <= 0.5)

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

    def process_freeformat_swel_entry(self, calendar, entry):
        """
        Process an entry that is a event with no fixed schedule.

        These events run continuously for several days, followed by
        separate semi-final and finals.  This is the same as an
        all-week-event, except that the event duration and name are wrong.
        """

        duration = self.meta.eventmeta[entry.code].duration
        duration = duration if duration else 51
        label = "%s R%s/%s" % (entry.name, 1, entry.rounds)
        self.process_all_week_entry(calendar, entry, duration, label)

    def process_freeformat_grognard_entry(self, calendar, entry):
        """
        Process an entry that is a pre-con event with no fixed schedule.

        These events run for 10 hours on Saturday, 15 hours on Sunday,
        15 hours on Monday, and 9 hours on Tuesday, before switching to a
        normal tourney schedule.  This is the same as an all-week-event,
        except that the event duration and name are wrong.

        In this case, the duration is 10 + 15 + 15 + 9 = 49 hours.
        """
        # FIXME: This is wrong for BWD, which starts at 10am on the PC days, not 9am
        duration = self.meta.eventmeta[entry.code].grognard
        duration = duration if duration else 49
        label = "%s R%s/%s" % (entry.name, 1, entry.rounds)
        self.process_all_week_entry(calendar, entry, duration, label)

    def process_all_week_entry(self, calendar, entry, length=None, label=None):
        """
        Process an entry that runs continuously all week long.
        """

        start = entry.datetime
        remaining = timedelta(hours=length) if length else entry.length
        label = label if label else entry.name + ' R1/1'

        while remaining.days or remaining.seconds:
            midnight = start.date() + timedelta(days=1)
            duration = datetime(midnight.year, midnight.month, midnight.day) - start
            if duration > remaining:
                duration = remaining

            self.add_event(calendar, entry, start=start, duration=duration, replace=False, name=label)

            start = datetime(midnight.year, midnight.month, midnight.day, 9, 0, 0)
            remaining = remaining - duration

    def alternate_round_name(self, entry, event_type=None):
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
        if entry.code in self.rounds and event_type in self.SINGLE:
            r = self.rounds[entry.code]
            offset = (len(self.SINGLE) - self.SINGLE.index(event_type)) - 1
            alternative = "%s R%s/%s" % (entry.name, r - offset, r)
        return alternative

