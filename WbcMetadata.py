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

import csv
import json
import logging
import os
import re
import sys
from collections import OrderedDict
from datetime import datetime
from optparse import OptionParser

from bs4 import NavigableString, Tag

from WbcUtility import parse_url, TZ, normalize, nu_strip

LOG = logging.getLogger('WbcMetaData')


# ----- JSON classes ----------------------------------------------------------

class WbcJsonEncoder(json.JSONEncoder):
    def default(self, o):
        if type(o) == WbcMetaEvent:
            return o.as_json()

        return json.JSONEncoder.default(self, o)


# ----- WBC Meta Data ---------------------------------------------------------

class WbcMetaEvent(object):
    def __init__(self, c, n):
        self.code = c
        self.name = n.strip()
        self.grognard = None
        self.duration = None
        self.playlate = None
        self.altnames = []

    def as_json(self):
        j = OrderedDict()
        j['name'] = self.name
        if self.duration:
            j['duration'] = self.duration
        if self.grognard:
            j['grognard'] = self.grognard
        if self.playlate:
            j['playlate'] = self.playlate
        altnames = set(self.altnames)
        altnames = altnames - set(self.name)
        altnames = list(altnames)
        altnames.sort()
        if len(altnames):
            j['altnames'] = altnames
        return j

    @classmethod
    def load_csv(cls, pathname):
        entries = OrderedDict()

        with open(pathname, 'r') as f:
            codefile = csv.DictReader(f, restkey='altnames')
            for row in codefile:
                code = row['Code'].strip()
                entry = WbcMetaEvent(code, row['Name'])
                entry.duration = int(row['Duration']) if row['Duration'] else None
                entry.grognard = int(row['Grognard']) if row['Grognard'] else None
                entry.playlate = row['PlayLate'].strip().lower() if row['PlayLate'] else None

                if row['altnames']:
                    entry.altnames = [x.strip() for x in row['altnames'] if x]

                entries[code] = entry

        return entries

    @classmethod
    def save_json(cls, pathname, entries):
        with open(pathname, 'w') as f:
            json.dump(entries, f, indent=2, cls=WbcJsonEncoder)

    @classmethod
    def load_json(cls, pathname):
        entries = OrderedDict()

        with open(pathname, 'r') as f:
            data = json.load(f)
            for code in sorted(data.keys()):
                row = data[code]
                entry = WbcMetaEvent(code, row['name'])
                entry.duration = row['duration'] if 'duration' in row else None
                entry.grognard = row['grognard'] if 'grognard' in row else None
                entry.playlate = row['playlate'] if 'playlate' in row else None
                entry.altnames = row['altnames'] if 'altnames' in row else []
                entries[code] = entry

        return entries


class WbcMetaOther(object):
    def __init__(self):
        pass

    @classmethod
    def load_csv(cls, pathname):
        entries = OrderedDict()

        return entries


class WbcMetadata(object):
    """Load metadata about events that is not available from other sources"""

    now = datetime.now(TZ)
    this_year = now.year
    yy = this_year % 100  # Yes, this will probably fail in 2100

    # List of events to debug
    tracking = [
        # 'EVL',
        # 'SSB',
        # 'WAW',
        # 'AFK', 'BWD', 'GBG', 'PZB', 'SQL', 'TRC', 'WAT', 'WSM'
    ]

    # # List of events for my personal calendar
    # personal = [
    #     '7WD', 'COB', 'C&K', 'IOV', 'KOT',
    #     'PGC', 'PGD', 'PRO', 'RFG', 'RGD', 'RRY',
    #     'SCY', 'SJN', 'SMW', 'SPD', 'TAM', 'TTR', 'T_M', 'TFM', 'VSD',
    # ]

    # Data file names
    EVENTCODES = os.path.join("meta", "wbc-event-codes.csv")
    OTHERCODES = os.path.join("meta", "wbc-other-codes.csv")

    SITE_URL = "http://boardgamers.org/wbc%02d/" % yy
    PREVIEW_INDEX_URL = SITE_URL + "previews-%d.html"

    # Bad code in event preview index -> actual event code {'gmb': 'GBM',}
    MISCODES = {
        '8xx': '8XX',
        'B-17': 'B17',
        'Iron Men - WSM': 'WSM',
    }

    SPECIAL_PREVIEWS = ['Juniors', 'Junior Events', 'Seminars', 'Demo', 'Demos', 'Demonstrations']

    others = []  # List of non-tournament event matching data
    special = []  # List of non-tournament event codes
    tourneys = []  # List of tournament codes

    codes = {}  # Name -> Code map for events
    names = {}  # Code -> Name map for events

    eventmeta = {}  # Meta data for tournament events
    othermeta = {}  # Meta data for non-tournament events

    durations = {}  # Special durations for events that have them
    grognards = {}  # Special durations for grognard events that have them
    playlate = {}  # Flag for events that may run past midnight
    url = {}  # Code -> URL for event preview for this event code

    first_day = None  # First calendar day for this year's convention
    last_day = None  # Last calendar day for this year's convention

    day_names = [
        'First Friday', 'First Saturday', 'First Sunday',
        'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday',
        'Second Monday'
    ]
    day_codes = ['FFr', 'FSa', 'FSu', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su', 'SMo']

    year = this_year  # Year to process
    type = 'new'  # Type of spreadsheet to parse
    input = None  # Name of spreadsheet to parse
    output = 'build'  # Name of directory for results
    write_files = True  # Whether or not to output files
    full_report = False  # True for more detailed discrepancy reporting
    verbose = False  # True for more detailed logging
    debug = False  # True for debugging and even more detailed logging

    def __init__(self):
        self.process_options()
        if self.type == 'old':
            self.load_tourney_codes()
            self.load_other_codes()
        self.load_preview_index()

    def process_options(self):
        """
        Parse command line options
        """

        parser = OptionParser()
        parser.add_option("-y", "--year", dest="year", metavar="YEAR", default=self.this_year, help="Year to process")
        parser.add_option("-t", "--type", dest="type", metavar="TYPE", default="new", help="Type of file to process (old,new)")
        parser.add_option("-i", "--input", dest="input", metavar="FILE", default=None, help="Schedule spreadsheet to process")
        parser.add_option("-o", "--output", dest="output", metavar="DIR", default="build", help="Directory for results")
        parser.add_option("-f", "--full-report", dest="fullreport", action="store_true", default=False)
        parser.add_option("-n", "--dry-run", dest="write_files", action="store_false", default=True)
        parser.add_option("-v", "--verbose", dest="verbose", action="store_true", default=False)
        parser.add_option("-d", "--debug", dest="debug", action="store_true", default=False)

        options, dummy_args = parser.parse_args()

        self.year = int(options.year)
        self.type = options.type
        self.input = options.input
        self.output = options.output
        self.full_report = options.fullreport
        self.write_files = options.write_files
        self.verbose = options.verbose
        self.debug = options.debug

        if self.debug:
            logging.root.setLevel(logging.DEBUG)
        elif self.verbose:
            logging.root.setLevel(logging.INFO)
        else:
            logging.root.setLevel(logging.WARN)

    def load_tourney_codes(self):
        """
        Load all of the tourney codes (and alternate names) from their data file.
        """

        LOG.debug('Loading tourney event codes')
        self.eventmeta = WbcMetaEvent.load_json(os.path.join('meta', 'wbc-event-codes.json'))

        for code, entry in self.eventmeta.items():
            self.codes[entry.name] = code
            self.names[code] = entry.name
            self.tourneys.append(code)
            for altname in entry.altnames:
                self.codes[altname] = code

    def load_other_codes(self):
        """
        Load all of the non-tourney codes from their data file.
        """

        LOG.debug('Loading non-tourney event codes')

        codefile = csv.DictReader(open(self.OTHERCODES))
        for row in codefile:
            c = row['Code'].strip()
            d = row['Description'].strip()
            n = row['Name'].strip()
            f = row['Format'].strip()

            other = {'code': c, 'description': d, 'name': n, 'format': f}
            self.others.append(other)
            self.special.append(c)
            self.names[c] = d

    def load_preview_index(self):
        """
        Load all of the links to the event previews and map them to event codes
        """

        LOG.debug('Loading event preview index')

        url = self.PREVIEW_INDEX_URL % self.year
        index = parse_url(url)
        if not index:
            LOG.error('Unable to load Preview index: %s', url)
            return

        # Find the preview table
        table = index.find('table').find('table').find('table').findAll('table')[1]
        rows = list(table.findAll('tr'))
        line = 0
        while line < len(rows):
            column = -1
            try:
                if len(rows) - line < 3:
                    LOG.error("Preview lines out of sync -- quitting")
                    continue

                top = list(rows[line].findAll('td'))
                mid = list(rows[line + 1].findAll('td'))

                if type(top[0].contents[0]) == NavigableString:
                    LOG.warn("Preview lines out of sync -- resyncing")
                    line += 1
                    continue

                for column in range(0, 8):  # Always 8 cells
                    link = name = code = None

                    top_column = top[column]
                    mid_column = mid[column]

                    if type(top_column.contents[0]) == NavigableString:
                        continue  # No more useful data on this row

                    link = top_column.a['href']

                    if link.endswith('wsm.html'):
                        LOG.debug('debug')
                        pass

                    mcc = mid_column.contents

                    mcc_text = [ nu_strip(e) for e in mcc if isinstance(e, NavigableString) ]
                    mcc_text = [ e for e in mcc_text if e ]

                    last = -1
                    if mcc_text[last].startswith('Under Construction'):
                        continue

                    if mcc_text[last].startswith('Updated'):
                        last = -2
                    
                    code = mcc_text[last]
                    name = ' '.join(mcc_text[:last])

                    if code in self.SPECIAL_PREVIEWS:
                        continue


                    # if len(mcc) == 5:
                    #     if type(mcc[4]) == NavigableString:
                    #         name = nu_strip(mcc[0]) + ' ' + nu_strip(mcc[2])
                    #         code = nu_strip(mcc[4])
                    #     else:
                    #         name = nu_strip(mcc[0])
                    #         code = nu_strip(mcc[2])
                    # elif isinstance(mcc[0], NavigableString):
                    #     name = nu_strip(mcc[0])
                    #     if name in self.SPECIAL_PREVIEWS:
                    #         LOG.info
                    #         continue  # FIXME: Grab URL for later use

                    #     m = re.match("(.*)-\s*(\w\w\w)", name)
                    #     if m:
                    #         name = normalize(str(m.group(1)))
                    #         code = normalize(str(m.group(2)))
                    #     else:
                    #         code = nu_strip(mcc[2])
                    # else:
                    #     if mid_column.p:
                    #         name = nu_strip(mid_column.p.contents[0])
                    #         if name in self.SPECIAL_PREVIEWS:
                    #             continue  # Not a single event schedule
                    #         code = nu_strip(mid_column.p.contents[2])
                    #     elif mid_column.span:
                    #         name = nu_strip(mid_column.span.contents[0])
                    #         if len(mid_column.span.contents) == 3:
                    #             code = nu_strip(mid_column.span.contents[2])
                    #         elif name in self.SPECIAL_PREVIEWS:
                    #             continue  # Not a single event schedule
                    #         else:
                    #             m = re.match("(.*) - (...)", name)
                    #             if m:
                    #                 name = m.group(1)
                    #                 code = m.group(2)
                    #             else:
                    #                 raise AssertionError("Cannot parse name/code")

                    # code = code.replace('(', '').replace(')', '')

                    # Map page codes to event codes
                    if code in self.MISCODES:
                        code = self.MISCODES[code]
                    else:
                        code = code.upper()

                    if link.startswith('http'):
                        self.url[code] = link
                    else:
                        self.url[code] = self.SITE_URL + link

            except Exception as e:
                LOG.exception("Skipping preview row %d, column %d:", line / 3, column)
                exc_type, exc_obj, exc_tb = sys.exc_info()
                LOG.warn("On line %d, skipping preview row %d, column %d: %s", exc_tb.tb_lineno, line / 3, column, getattr(e, 'message', '---'))
                pass

            line += 3  # Lines should be in groups of 3

        LOG.warn("Found %d events in preview", len(self.url))

    def check_date(self, event_date):
        """Check to see if this event date is the earliest event date seen so far"""
        if not event_date:
            pass

        if not self.first_day:
            self.first_day = event_date
        elif event_date < self.first_day:
            self.first_day = event_date

        if not self.last_day:
            self.last_day = event_date
        elif event_date > self.last_day:
            self.last_day = event_date
