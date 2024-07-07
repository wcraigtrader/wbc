#! /usr/bin/env python3

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

"""WbcCalendars: Generate iCal calendars from the WBC Schedule spreadsheet"""

# xxlint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0612,W0621,W0702,W0703
# pylint: disable=C0103,C0301,C0302,R0902,R0903,R0904,R0912,R0913,R0914,W0702

import codecs
import csv
import json
import logging
import os
import shutil
import urllib.request
import urllib.parse
import urllib.error
import warnings

from datetime import date
from functools import cmp_to_key
from operator import attrgetter

from bs4 import BeautifulSoup, GuessedAtParserWarning
from icalendar import Calendar, Event

from WbcUtility import as_local, cmp, globalize, round_up_datetime


warnings.filterwarnings('ignore', category=GuessedAtParserWarning)

LOG = logging.getLogger('WbcCalendars')


class WbcWebcal(object):
    calendars = {}  # Calendars for each event code
    locations = {}  # Calendars by location
    dailies = {}  # Calendars by date

    current_tourneys = None
    everything = None
    tournaments = None
    meta = None

    # Data file names
    TEMPLATE = "resources/index-template.html"
    ICONS = ['ical.png', 'gcal16.png']
    # KEYS =  lambda e:

    def __init__(self, metadata):
        """
        Initialize the iCal calendars
        """

        self.meta = metadata

        self.prodid = "WBC %s" % self.meta.year

        if not os.path.exists(self.meta.output):
            os.makedirs(self.meta.output)

    def create_event(self, calendar, entry, start=None, duration=None, name=None, altname=None, replace=True):
        """
        Add a new vEvent to the given iCalendar for a given spreadsheet entry.
        """
        try:
            name = name if name else entry.name
            start = start if start else entry.datetime
            start = round_up_datetime(start)
            duration = duration if duration else entry.length

            url = self.meta.url[entry.code] if entry.code in self.meta.url else ''

            description = name
            if entry.code:
                description = entry.code + ': ' + description
            if entry.rtype:
                description += ' ' + entry.rtype
            if entry.format:
                description += ' (' + entry.format + ')'
            if entry.continuous:
                description += ' Continuous'
            if url:
                description += '\nPreview: ' + urllib.parse.quote(url, ':/')

            e = Event()
            e.add('SUMMARY', name)
            e.add('DESCRIPTION', description)
            e.add('DTSTART', globalize(start))
            e.add('DURATION', duration)
            e.add('LOCATION', entry.location)
            e.add('CONTACT', entry.gm)
            e.add('URL', url)
            e.add('LAST-MODIFIED', self.meta.now)
            e.add('DTSTAMP', self.meta.now)
            e.add('UID', f"{self.prodid}: {name}")
            e.add('COMMENT', repr(entry.extra))

            if replace:
                self.add_or_replace_event(calendar, e, altname)
            else:
                calendar.add_component(e)

        except TypeError as e:
            pass

    def create_all_calendars(self):

        # Create calendar events from all of the spreadsheet events.
        self.create_wbc_calendars()

    def create_wbc_calendars(self):
        """
        Process all of the spreadsheet entries, by event code, then by time,
        creating calendars for each entry as needed.
        """

        LOG.info('Creating calendars')

        # Create a sorted list of this year's tourney codes
        self.current_tourneys = [code for code in self.calendars.keys() if code in self.meta.tourneys]
        self.current_tourneys.sort(key=lambda k: self.meta.names[k])

        # Create bulk calendars
        self.everything = Calendar()
        self.everything.add('VERSION', '2.0')
        self.everything.add('PRODID', '-//' + self.prodid + ' Everything//ct7//')
        self.everything.add('SUMMARY', 'WBC %s All-in-One Schedule' % self.meta.year)

        self.tournaments = Calendar()
        self.tournaments.add('VERSION', '2.0')
        self.tournaments.add('PRODID', '-//' + self.prodid + ' Tournaments//ct7//')
        self.tournaments.add('SUMMARY', 'WBC %s Tournaments Schedule' % self.meta.year)

        # For all of the event calendars
        for code, calendar in self.calendars.items():

            # Add all calendar events to the master calendar
            self.everything.subcomponents += calendar.subcomponents

            # Add all the tourney events to the tourney calendar
            if code in self.current_tourneys:
                self.tournaments.subcomponents += calendar.subcomponents

            # For each calendar event
            for event in calendar.subcomponents:
                # Add it to the appropriate location calendar
                location = self.get_or_create_location_calendar(event['LOCATION'])
                location.subcomponents.append(event)

                # Add it to the appropriate daily calendar
                daily = self.get_or_create_daily_calendar(event['DTSTART'])
                daily.subcomponents.append(event)

    def get_or_create_event_calendar(self, code):
        """
        For a given event code, return the iCalendar that matches that code.
        If there is no pre-existing calendar, create a new one.
        """

        if code in ['Demo', 'Demonstrations']:
            code = 'Demos'

        if code in self.calendars:
            return self.calendars[code]

        if code not in self.meta.names:
            pass
        
        description = "%s %s: %s" % (self.prodid, code, self.meta.names[code])

        calendar = Calendar()
        calendar.add('VERSION', '2.0')
        calendar.add('PRODID', '-//%s %s//ct7//' % (self.prodid, code))
        calendar.add('SUMMARY', self.meta.names[code])
        calendar.add('DESCRIPTION', description)
        if code in self.meta.url:
            calendar.add('URL', self.meta.url[code])

        self.calendars[code] = calendar

        return calendar

    def get_or_create_location_calendar(self, location):
        """
        For a given location, return the iCalendar that matches that location.
        If there is no pre-existing calendar, create a new one.
        """
        location = str(location).strip()
        if location in self.locations:
            return self.locations[location]

        description = "%s: Events in %s" % (self.prodid, location)

        calendar = Calendar()
        calendar.add('VERSION', '2.0')
        calendar.add('PRODID', '-//%s %s//ct7//' % (self.prodid, location))
        calendar.add('SUMMARY', 'Events in ' + location)
        calendar.add('DESCRIPTION', description)

        self.locations[location] = calendar

        return calendar

    def get_or_create_daily_calendar(self, event_date):
        """
        For a given date, return the iCalendar that matches that date.
        If there is no pre-existing calendar, create a new one.
        """
        key = event_date.dt.date()
        name = event_date.dt.strftime('%A, %B %d')

        if key in self.dailies:
            return self.dailies[key]

        description = '%s: Events on %s' % (self.prodid, name)

        calendar = Calendar()
        calendar.add('VERSION', '2.0')
        calendar.add('PRODID', '-//%s %s//ct7//' % (self.prodid, key))
        calendar.add('SUMMARY', 'Events on ' + name)
        calendar.add('DESCRIPTION', description)

        self.dailies[key] = calendar

        return calendar

    def write_calendar_file(self, calendar, name):
        """
        Write an actual calendar file, using a filesystem-safe name.
        """
        filename = self.safe_ics_filename(name)
        with open(os.path.join(self.meta.output, filename), "wb") as f:
            f.write(self.serialize_calendar(calendar))

    def write_all_calendar_files(self):
        """
        Write all of the calendar files.
        """
        LOG.info("Saving calendars...")

        # For all of the event calendars
        for code, calendar in self.calendars.items():
            # Write the calendar itself
            self.write_calendar_file(calendar, code)

        # Write the master and tourney calendars
        self.write_calendar_file(self.everything, "all-in-one")
        self.write_calendar_file(self.tournaments, "tournaments")

        # Write the location calendars
        for location, calendar in self.locations.items():
            self.write_calendar_file(calendar, location)

        # Write the daily calendars
        for day, calendar in self.dailies.items():
            self.write_calendar_file(calendar, day)

    def write_json_files(self):
        """
        Write all of the calendar entries back out, in JSON format, with improvements
        """
        LOG.info('Writing JSON spreadsheet...')

        json_filename = os.path.join(self.meta.output, "details.json")

        with codecs.open(json_filename, 'w', 'utf-8') as json_file:
            json_file.write('[\n')
            subsequent = False
            for event in self.everything.subcomponents:
                json_file.write(',\n  ' if subsequent else '  ')
                subsequent = True

                row = eval(event['COMMENT'])
                row['Continuous'] = 'Y' if row['Continuous'] else ''
                row['Event'] = event['SUMMARY']
                row['GM'] = event['CONTACT']
                row['Location'] = event['LOCATION']
                sdatetime = as_local(event.decoded('DTSTART'))
                row['Date'] = sdatetime.date().strftime('%m/%d/%Y')
                stime = sdatetime.time()
                xtime = stime.hour * 1.0 + stime.minute / 60.0
                xtime = int(xtime) if int(xtime) == xtime else xtime
                row['Time'] = xtime
                duration = event.decoded('DURATION')
                row['Duration'] = duration.total_seconds() / 3600.0

                json_file.write(json.dumps(row))

            json_file.write('\n]\n')

    def write_csv_details(self, header):
        """
        Write all of the calendar entries back out, in CSV format, with improvements
        """
        LOG.info('Writing CSV spreadsheet...')

        details_filename = os.path.join(self.meta.output, "details.csv")

        with codecs.open(details_filename, "w", 'utf-8') as csv_file:
            writer = csv.DictWriter(csv_file, header, extrasaction='ignore')
            writer.writeheader()
            for event in self.everything.subcomponents:
                row = eval(event['COMMENT'])
                row['Continuous'] = 'Y' if row['Continuous'] else ''
                row['Event'] = event['SUMMARY']
                row['GM'] = event['CONTACT']
                row['Location'] = event['LOCATION']
                sdatetime = as_local(event.decoded('DTSTART'))
                row['Date'] = sdatetime.date().strftime('%Y-%m-%d')
                stime = sdatetime.time()
                row['Time'] = stime.hour * 1.0 + stime.minute / 60.0
                duration = event.decoded('DURATION')
                row['Duration'] = duration.total_seconds() / 3600.0
                writer.writerow(row)

    def write_index_page(self):
        """
        Using an HTML Template, create an index page that lists
        all of the created calendars.
        """

        LOG.info('Writing index page...')

        # Copy needed files to the destination
        for filename in self.ICONS:
            source = os.path.join("resources", filename)
            if os.path.exists(source):
                shutil.copy(source, self.meta.output)

        with open(self.TEMPLATE, "r") as f:
            template = f.read()

        parser = BeautifulSoup(template, 'lxml')

        # Locate insertion points
        title = parser.find('title')
        header = parser.find('div', {'id': 'header'})
        footer = parser.find('div', {'id': 'footer'})

        # Page title
        title.insert(0, parser.new_string("WBC %s Event Schedule" % self.meta.year))
        header.h1.insert(0, parser.new_string("WBC %s Event Schedule" % self.meta.year))
        footer.p.insert(0, parser.new_string("Updated on %s" % self.meta.now.strftime("%A, %d %B %Y %H:%M %Z")))

        # Tournament event calendars
        tourneys = dict([(k, v) for k, v in self.calendars.items() if k not in self.meta.special])
        self.render_calendar_table(parser, 'tournaments', 'Tournament Events', tourneys, lambda k: tourneys[k]['summary'])

        # Non-tourney event calendars
        nontourneys = dict([(k, v) for k, v in self.calendars.items() if k in self.meta.special])
        self.render_calendar_list(parser, 'other', 'Other Events', nontourneys)

        # Location calendars
        self.render_calendar_list(parser, 'location', 'Location Calendars', self.locations)

        # Daily calendars
        self.render_calendar_list(parser, 'daily', 'Daily Calendars', self.dailies)

        # Special event calendars
        specials = {
            'all-in-one': self.everything,
            'tournaments': self.tournaments,
        }
        self.render_calendar_list(parser, 'special', 'Special Calendars', specials)

        with codecs.open(os.path.join(self.meta.output, 'index.html'), 'w', 'utf-8') as f:
            f.write(parser.prettify())

    @classmethod
    def render_calendar_table(cls, parser, id_name, label, calendar_map, key=None):
        """Create the HTML fragment for the table of tournament calendars."""

        keys = list(calendar_map.keys())
        keys.sort(key=key)

        div = parser.find('div', {'id': id_name})
        div.insert(0, parser.new_tag('h2'))
        div.h2.insert(0, parser.new_string(label))
        div.insert(1, parser.new_tag('table'))

        for row_keys in cls.split_list(keys, 2):
            tr = parser.new_tag('tr')
            div.table.insert(len(div.table), tr)

            for key in row_keys:
                label = calendar_map[key]['summary'] if key else ''
                td = cls.render_calendar_table_entry(parser, key, label)
                tr.insert(len(tr), td)

    @classmethod
    def render_calendar_table_entry(cls, parser, key, label):
        """Create the HTML fragment for one cell in the tournament calendar table."""
        td = parser.new_tag('td')
        if key:
            span = parser.new_tag('span')
            span['class'] = 'eventcode'
            span.insert(0, parser.new_string(key + ': '))
            td.insert(len(td), span)

            filename = cls.safe_ics_filename(key)

            a = parser.new_tag('a')
            a['class'] = 'eventlink'
            a['href'] = '#'
            a['onclick'] = "webcal('%s');" % filename
            img = parser.new_tag('img')
            img['src'] = cls.ICONS[0]
            a.insert(len(a), img)
            td.insert(len(td), a)

            a = parser.new_tag('a')
            a['class'] = 'eventlink'
            a['href'] = '#'
            a['onclick'] = "gcal('%s');" % filename
            img = parser.new_tag('img')
            img['src'] = cls.ICONS[1]
            a.insert(len(a), img)
            td.insert(len(td), a)

            a = parser.new_tag('a')
            a['class'] = 'eventlink'
            a['href'] = filename
            a.insert(0, parser.new_string("%s" % label))
            td.insert(len(td), a)

            td.insert(len(td), a)
        else:
            td.insert(len(td), parser.new_string(' '))
        return td

    @staticmethod
    def split_list(original, width):
        """
        A generator that, given a list of indeterminate length, will split the list into
        roughly equal columns, and then return the resulting list one row at a time.
        """

        max_length = len(original)
        length = int((max_length + width - 1) / width)
        for i in range(length):
            partial = []
            for j in range(width):
                k = i + j * length
                partial.append(original[k] if k < max_length else None)
            yield partial

    @classmethod
    def render_calendar_list(cls, parser, id_name, label, calendar_map, key=None):
        """Create the HTML fragment for an unordered list of calendars."""

        keys = list(calendar_map.keys())
        keys.sort(key=key)

        div = parser.find('div', {'id': id_name})
        div.insert(0, parser.new_tag('h2'))
        div.h2.insert(0, parser.new_string(label))
        div.insert(1, parser.new_tag('ul'))

        for key in keys:
            calendar = calendar_map[key]
            cls.render_calendar_list_item(parser, div.ul, key, calendar['summary'])

    @classmethod
    def render_calendar_list_item(cls, parser, list_tag, key, label):
        """Create the HTML fragment for a single calendar in a list"""

        li = parser.new_tag('li')

        filename = cls.safe_ics_filename(key)

        a = parser.new_tag('a')
        a['class'] = 'eventlink'
        a['href'] = '#'
        a['onclick'] = "webcal('%s');" % filename
        img = parser.new_tag('img')
        img['src'] = cls.ICONS[0]
        a.insert(len(a), img)
        li.insert(len(li), a)

        a = parser.new_tag('a')
        a['class'] = 'eventlink'
        a['href'] = '#'
        a['onclick'] = "gcal('%s');" % filename
        img = parser.new_tag('img')
        img['src'] = cls.ICONS[1]
        a.insert(len(a), img)
        li.insert(len(li), a)

        a = parser.new_tag('a')
        a['class'] = 'eventlink'
        a['href'] = filename
        a.insert(0, parser.new_string("%s" % label))
        li.insert(len(li), a)

        list_tag.insert(len(list_tag), li)

    @classmethod
    def serialize_calendar(cls, calendar):
        """This fixes portability quirks in the iCalendar library:
        1) The iCalendar library generates event start date/times as 'DTSTART;DATE=VALUE:yyyymmddThhmmssZ';
           the more acceptable format is 'DTSTART:yyyymmddThhmmssZ'
        2) The iCalendar library doesn't sort the events in a given calendar by date/time.
        """

        c = calendar
        c.subcomponents.sort(key=cmp_to_key(cls.compare_icalendar_events))

        output = c.to_ical()
        # output = output.replace( ";VALUE=DATE-TIME:", ":" )
        return output

    @classmethod
    def add_or_replace_event(cls, calendar, event, altname=None):
        """
        Insert a vEvent into an iCalendar.
        If the vEvent 'matches' an existing vEvent, replace the existing vEvent instead.
        """
        for i in range(len(calendar.subcomponents)):
            if cls.is_same_icalendar_event(calendar.subcomponents[i], event, altname):
                calendar.subcomponents[i] = event
                return
        calendar.subcomponents.append(event)

    @classmethod
    def is_same_icalendar_event(cls, e1, e2, altname=None):
        """
        Compare two events to determine if they are 'the same'.

        If they start at the same time, and have the same duration, they're 'the same'.
        If they have the same name, they're 'the same'.
        If the first matches an alternative name, they're 'the same'.
        """
        same = str(e1['dtstart']) == str(e2['dtstart'])
        same &= str(e1['duration']) == str(e2['duration'])
        same |= str(e1['summary']) == str(e2['summary'])
        if altname:
            same |= str(e1['summary']) == str(altname)
        return same

    @staticmethod
    def compare_icalendar_events(x, y):
        """
        Comparison method for iCal events
        """
        c = cmp(x['dtstart'].dt, y['dtstart'].dt)
        c = cmp(x['summary'], y['summary']) if not c else c
        return c

    @staticmethod
    def safe_ics_filename(name):
        """
        Given an object, determine a web-safe filename from it, then append '.ics'.
        """
        if name.__class__ is date:
            name = name.strftime("%Y-%m-%d")
        else:
            name = name.strip()
            name = name.replace('&', 'n')
            name = name.replace(' ', '_')
            name = name.replace('/', '_')
        return "%s.ics" % name


if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)
    logging.getLogger('requests').setLevel(logging.WARN)
