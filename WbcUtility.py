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

from web import Web
from datetime import timedelta, datetime

import logging
import pytz
import unicodedata
import xlrd

LOGGER = logging.getLogger('WbcUtility')


# ----- Text Functions --------------------------------------------------------

def normalize(utext):
    return unicodedata.normalize('NFKD', utext).encode('ascii', 'ignore')


def nu_strip(string):
    return normalize(str(string)).strip()


# ----- Spreadsheet Functions -------------------------------------------------

def parse_value(data):
    if data.ctype == xlrd.XL_CELL_EMPTY:
        data = None
    elif data.ctype == xlrd.XL_CELL_TEXT:
        data = unicodedata.normalize('NFKD', data.value).encode('ascii', 'ignore').strip()
        if data == '--':  # Assume that '--' really means an empty cell
            data = None
    elif data.ctype == xlrd.XL_CELL_NUMBER:
        data = float(data.value)
    elif data.ctype == xlrd.XL_CELL_DATE:
        data = xlrd.xldate_as_tuple(data.value, 0)  # sheet.book.datemode
        if data[0]:
            data = datetime(*data)  # pylint: disable=W0142
        else:
            data = time(data[3], data[4], data[5])
    else:
        raise ValueError("Unhandled Excel cell type (%s)" % data.ctype)

    return data


def sheet_value(sheet, row, col):
    try:
        v = sheet.cell(row, col)
        return parse_value(v)
    except ValueError as e:
        raise ValueError(e.message + "@(%d, %d)" % (row, col))


# ----- Time Functions --------------------------------------------------------

TZ = pytz.timezone('America/New_York')  # Tournament timezone
UTC = pytz.timezone('UTC')  # UTC timezone (for iCal)


def as_local(timestamp):
    """Return the zoned timestamp, assuming local timezone"""
    return timestamp.astimezone(TZ)


def as_global(timestamp):
    """Return the zoned timestamp, assuming UTC timezone"""
    return timestamp.astimezone(UTC)


def localize(timestamp):
    """Return the unzoned timestamp, as a zoned timestamp, assuming local timezone"""
    return TZ.localize(timestamp)


def globalize(timestamp):
    """Return the unzoned timestamp, as a zoned timestamp, assuming UTC timezone"""
    return TZ.localize(timestamp).astimezone(UTC)


def round_up_datetime(timestamp):
    """Round up the timestamp to the nearest minute"""
    ts = timestamp + timedelta(seconds=30)
    ts = ts.replace(second=0, microsecond=0)
    return ts


def round_up_timedelta(duration):
    """Round up the duration to the nearest minute"""
    seconds = duration.total_seconds()
    seconds = 60 * int((seconds + 30) / 60)
    return timedelta(seconds=seconds)


def cal_time(timestamp):
    return as_local(round_up_datetime(timestamp))


def text_to_date(text):
    value = None
    try:
        if isinstance(value, str):
            value = datetime.strptime(text, '%m/%d/%y')
        if isinstance(value, datetime):
            value = value.date()
    except ValueError:
        pass  # Not a date, that's OK
    except TypeError:
        pass  # Not a string?
    return value


def text_to_datetime(text):
    value = None
    try:
        if isinstance(text, datetime):
            value = text
        elif isinstance(text, str):
            value = datetime.strptime(text, '%m/%d/%y')
    except ValueError as e:
        pass  # Not a date, that's OK
    return value


# ----- Web methods -----------------------------------------------------------


def parse_url(url):
    return Web.fetch(url)
