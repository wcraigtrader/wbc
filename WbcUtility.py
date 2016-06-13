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

from web import Web
from datetime import timedelta
import logging
import pytz
import unicodedata

LOGGER = logging.getLogger( 'WbcUtility' )

# ----- Text Functions --------------------------------------------------------

def normalize( utext ):
    return unicodedata.normalize( 'NFKD', utext ).encode( 'ascii', 'ignore' )

# ----- Time Functions --------------------------------------------------------

TZ = pytz.timezone( 'America/New_York' )  # Tournament timezone
UTC = pytz.timezone( 'UTC' )  # UTC timezone (for iCal)

def as_local(timestamp):
    """Return the zoned timestamp, assuming local timezone"""
    return timestamp.astimezone( TZ )

def as_global(timestamp):
    """Return the zoned timestamp, assuming UTC timezone"""
    return timestamp.astimezone( UTC )

def localize( timestamp ):
    """Return the unzoned timestamp, as a zoned timestamp, assuming local timezone"""
    return TZ.localize( timestamp )

def globalize( timestamp ):
    """Return the unzoned timestamp, as a zoned timestamp, assuming UTC timezone"""
    return TZ.localize( timestamp ).astimezone( UTC )

def round_up_datetime( timestamp ):
    """Round up the timestamp to the nearest minute"""
    ts = timestamp + timedelta( seconds=30 )
    ts = ts.replace( second=0, microsecond=0 )
    return ts

def round_up_timedelta( duration ):
    """Round up the duration to the nearest minute"""
    seconds = duration.total_seconds()
    seconds = 60 * int( ( seconds + 30 ) / 60 )
    return timedelta( seconds=seconds )

def cal_time( timestamp ):
    return as_local( round_up_datetime ( timestamp ) )

# ----- Web methods -----------------------------------------------------------

def parse_url( url):
    return Web.fetch( url )
