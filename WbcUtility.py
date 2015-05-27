#----- Copyright (c) 2010-2015 by W. Craig Trader ---------------------------------
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

from bs4 import BeautifulSoup
from hashlib import md5 as hash
import logging
import os
import urllib2

LOGGER = logging.getLogger( 'WbcUtility' )

#----- Utility methods -------------------------------------------------------

def parse_url( url ):
    """
    Utility function to load an HTML page from a URL, and parse it with BeautifulSoup.
    """

    WEBCACHE = 'cache'

    if not os.path.exists( WEBCACHE ):
        os.makedirs( WEBCACHE )


    page = None
    try:
        cacheid = os.path.join( WEBCACHE, hash( url ).hexdigest() )
        if os.path.exists( cacheid ):
            with open( cacheid, 'r' ) as c:
                data = c.read()
        else:
            f = urllib2.urlopen( url )
            data = f.read()
            with open( cacheid, 'w' ) as c:
                c.write( data )

        if ( len( data ) ):
            page = BeautifulSoup( data, "lxml" )

    except Exception as e:  # pylint: disable=W0703
        LOGGER.error( 'Failed while loading (%s)', url )
        LOGGER.error( e )

    return page

