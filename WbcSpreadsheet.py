# ----- Copyright (c) 2010-2022 by W. Craig Trader ---------------------------------
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

import codecs
import csv
import logging
import os
import shutil
from collections import OrderedDict
from datetime import datetime
from functools import total_ordering

if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)
    logging.getLogger('requests').setLevel(logging.WARN)

LOG = logging.getLogger('WbcSpreadsheet')


# ----- WBC Row ---------------------------------------------------------------

@total_ordering
class WbcRow(object):
    """
    A WbcRow encapsulates information about a single schedule line from the
    WBC schedule spreadsheet. This line may result in a dozen or more calendar events.

    Date: MM/DD/YYYY
    Time: Integer, Event start hour
    Event: Event name with codes
    Prize: Number of plaques awarded (1-6)
    Class: 'A', 'B', 'C', otherwise blank if inapplicable
    Format: Tournament format for this session
    LN: 'Y' if late night event, otherwise blank
    FF: 'Y' if free form event, otherwise blank
    Continuous: 'Y' if this entry refers to multiple events, scheduled back-to-back
    GM: Game Master for this event
    Location: Which room name (and table name for demos)
    """

    def __init__(self, schedule, line, *args):

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
        self.location = None
        self.prize = None
        self.category = None

        self.type = ''
        self.rtype = ''
        self.rounds = 0
        self.freeformat = False
        self.grognard = False
        self.junior = False

        self.event = None

        # read the data row using the subclass
        self.readrow(*args)

        if self.gm is None:
            LOG.warning('Missing gm on %s', self)
            self.gm = ''

        if self.duration is None:
            LOG.warning('Missing duration on %s', self)
            self.duration = '0'

        self.initialize()

        if not isinstance(self.date, datetime):
            LOG.error('Unreadable date on %s', self)
            raise ValueError('Unreadable date (%s, %s)' % (type(self.date), self.date))

    @property
    def __key__(self):
        return self.code, self.datetime, self.name

    def __eq__(self, other):
        return self.__key__ == other.__key__

    def __lt__(self, other):
        return self.__key__ < other.__key__

    def __setattr__(self, key, value):
        k = key.strip().lower().replace(' ', '_')
        self.__dict__[k] = value

    def __getattr__(self, key):
        k = key.strip().lower()
        if k in self.__dict__:
            return self.__dict__[k]
        else:
            raise AttributeError(key)

    def __repr__(self):
        if isinstance(self.date, datetime):
            return "%s @ %s %s on %s" % (self.event, self.date.date(), self.time, self.line)
        else:
            return "%s @ %s %s on %s" % (self.event, self.date, self.time, self.line)

    @property
    def row(self):
        try:
            row = dict([(k, getattr(self, k)) for k in self.FIELDS])
            row['Date'] = self.date.strftime('%Y-%m-%d')
            row['Continuous'] = 'Y' if row['Continuous'] else ''
            event = self.name
            row['Event'] = event
        except Exception as e:
            LOG.error("Unexpected row exception: %s", e)
        return row

    @property
    def extra(self):
        extra = dict([(k, getattr(self, k)) for k in ['Code', 'Prize', 'Class', 'Format', 'Continuous'] if hasattr(self, k)])
        return extra

    def readrow(self, *args):
        """Stub"""

        raise NotImplementedError()

    def initialize(self):
        """Stub"""

        raise NotImplementedError()


# ----- WBC Schedule ----------------------------------------------------------

class WbcSchedule(object):
    """
    The WbcSchedule class parses the entire WBC schedule spreadsheet and creates
    iCalendar calendars for each event (with vEvents for each time slot).
    """
    valid = False

    tracking = []

    header = []
    output = ['Date', 'Time', 'Event', 'Prize', 'Class', 'Format', 'Duration', 'Continuous', 'GM', 'Location', 'Code']

    rounds = {}  # Number of rounds for events that have rounds
    events = {}  # Events, grouped by code and then sorted by start date/time
    unmatched = []  # List of spreadsheet rows that don't match any coded events

    meta = None
    calendars = None

    def __init__(self, metadata, calendars):
        """
        Initialize a schedule
        """

        self.meta = metadata
        self.calendars = calendars
        self.events = OrderedDict()

        self.filename = self.meta.input

        self.tracking.extend(metadata.tracking)
        self.load_events()

        LOG.info('First day is %s', self.meta.first_day.date())
        LOG.info('Last day is  %s', self.meta.last_day.date())

        self.report_unprocessed_events()

        self.valid = True

    def load_events(self):
        """Stub"""

        raise NotImplementedError()

    def categorize_row(self, row):
        """
        Assign a spreadsheet entry to a matching row code.
        """
        if row.code:
            if row.code in self.events:
                self.events[row.code].append(row)
            else:
                self.events[row.code] = [row]
        else:
            self.unmatched.append(row)

        if row.code in self.meta.tourneys:
            self.meta.check_date(row.date)

    def add_event(self, calendar, entry, start=None, duration=None, name=None, altname=None, replace=True):
        self.calendars.create_event(calendar, entry, start, duration, name, altname, replace)

    def report_unprocessed_events(self):
        """
        Report on all of the WBC schedule entries that were not processed.
        """

        # self.unmatched.sort(cmp=lambda x, y: cmp(x.name, y.name))
        for event in self.unmatched:
            LOG.error('Did not process Row %5d [%s] %s', event.line, event.name, event)

    def write_all_files(self):
        # Remote any existing destination directory
        if os.path.exists(self.meta.output):
            shutil.rmtree(self.meta.output)

        # Create the destination directory
        os.makedirs(self.meta.output)

        self.write_csv_spreadsheet()
        self.calendars.write_csv_details(self.output)
        self.calendars.write_json_files()
        self.calendars.write_index_page()
        self.calendars.write_all_calendar_files()

    def write_csv_spreadsheet(self):
        """
        Write all of the calendar entries back out, in CSV format, with improvements
        """
        LOG.info('Writing CSV spreadsheet...')

        # data = []
        # for k in self.events.keys():
        #     data = data + self.events[k]
        # data = data + self.unmatched
        # data.sort()

        spreadsheet_filename = os.path.join(self.meta.output, "schedule.csv")

        # FIXME: Polymorphic FIELDS

        with codecs.open(spreadsheet_filename, "w", 'utf-8') as csv_file:
            writer = csv.DictWriter(csv_file, self.output, extrasaction='ignore')
            writer.writeheader()
            for event_list in self.events.values():
                for entry in event_list:
                    writer.writerow(entry.row)
