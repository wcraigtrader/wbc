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

