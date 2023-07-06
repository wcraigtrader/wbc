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

from datetime import timedelta
import logging

from WbcUtility import parse_url, localize, cmp


LOGGER = logging.getLogger('WbcAllInOne')


# ----- WBC All-in-One Schedule -----------------------------------------------


class WbcAllInOne(object):
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
        &nbsp</tr>

    This is, frankly, horrible HTML.  But at least it's consistent, year-to-year,
    and BeautifulSoup can parse it.
    """

    SITE_URL = 'http://boardgamers.org/wbc%d/allin1.htm'

    valid = False

    colormap = {
        'green': 'Demo',
        'magenta': 'Mulligan',
        'red': 'Round',
        '#07BED2': 'QF',
        'blue': 'SF',
        '#AAAA00': 'F',
    }

    # Events that are miscoded (bad code : actual code)
    # codemap = { 'MMA': 'MRA', }
    codemap = {'T_G': 'T-G'}
    roommap = {
        'Festival': 'Festival Hall',
        'Ballroom B': 'Ballroom',
        'First Tracks Pool': 'First Tracks Poolside',
        'First TracksPoolside': 'First Tracks Poolside',
    }
    events = {}

    notes = {
        '7WD': 'All-in-One does not handle 45 minute rounds',
        'CNS': 'All-in-One does not handle 30 minute rounds',
        'ELC': 'All-in-One does not handle 20 minute rounds',
        'KOT': 'All-in-One does not handle 45 minute rounds',
        'LID': 'All-in-One does not handle 30 minute rounds',
    }

    class Event(object):
        """Simple data object to collect information about an event occuring at a specific time."""

        def __init__(self):
            self.code = None
            self.name = None
            self.type = None
            self.time = None
            self.location = None

        def __cmp__(self, other):
            return cmp(self.time, other.time)

        def __str__(self):
            return '%s %s %s in %s at %s' % (self.code, self.name, self.type, self.location, self.time)

    def __init__(self, metadata):
        self.meta = metadata
        self.page = None

        self.load_table()

    def load_table(self):
        """Parse the All-in-One schedule (HTML)"""

        LOGGER.info('Parsing WBC All-in-One schedule')

        self.page = parse_url(self.SITE_URL % (self.meta.year % 100))
        if not self.page:
            return

        try:
            title = self.page.findAll('title')[0]
            year = str(title.text)
            year = year.strip().split()
            year = int(year[0])
        except:
            # Fetch from page body instead of page title.
            # html.body.table.tr.td.p.font.b.font.NavigableString
            try:
                td = self.page.html.body.table.tr.td
                text = td.h1.b.text
                year = str(text).strip().split()
                year = int(year[0])
            except:
                year = 2013

        if year != self.meta.this_year and year != self.meta.year:
            LOGGER.error("All-in-one schedule for %d is out of date", year)

            return

        tables = self.page.findAll('table')
        rows = tables[1].findAll('tr')
        for row in rows[1:]:
            self.load_row(row)

        self.valid = True

    def load_row(self, row):
        """Parse an individual all-in-one row to find times and rooms for an event"""

        events = []

        cells = row.findAll('td')
        code = str(cells[0].font.text).strip(';')
        name = str(cells[1].font.text).strip(';* ')

        code = self.codemap[code] if code in self.codemap else code

        current_date = self.meta.first_day

        # For each day ...
        for cell in cells[3:]:
            current = {}

            # All entries belong to font tags
            for f in cell.findAll('font'):
                for key, val in f.attrs.items():
                    if key == 'color':
                        # Fonts with color attributes represent start/type data for a single event
                        e = WbcAllInOne.Event()
                        e.code = code
                        e.name = name
                        hour = int(f.text.strip())
                        day = current_date.day
                        month = current_date.month
                        if hour >= 24:
                            hour -= 24
                            day += 1
                        if day >= 32:  # This works because WBC always starts in either the end of July or beginning of August
                            day -= 31
                            month += 1
                        e.time = localize(current_date.replace(month=month, day=day, hour=hour))
                        e.type = self.colormap.get(val, None)
                        current[hour] = e

                    elif key == 'size':
                        # Fonts with size=-1 represent entry data for all events
                        text = str(f.text).strip().split('; ')

                        if len(text) == 1:
                            # If there's only one entry, it applies to all events
                            entry = text[0]
                            entry = self.roommap[entry] if entry in self.roommap else entry
                            for e in current.values():
                                e.location = entry
                        else:
                            # For each entry ...
                            for chunk in text:
                                times, dummy, entry = chunk.partition(':')
                                entry = self.roommap[entry] if entry in self.roommap else entry
                                if times == 'others':
                                    # Apply this location to all entries without locations
                                    for e in current.values():
                                        if not e.location:
                                            e.location = entry
                                else:
                                    # Apply this location to each listed hour
                                    for hour in times.split(','):
                                        current[int(hour)].location = entry

            # Add all of this days events to the list
            events = events + list(current.values())

            # Move to the next date
            current_date = current_date + timedelta(days=1)

        # Sort the list, then add it to the events map
        events.sort()
        self.events[code] = events
