#! /usr/bin/env python2.7

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

import codecs
import logging
import time
from datetime import timedelta

import requests
import requests_cache
from bs4 import BeautifulSoup

# ----- Globals ---------------------------------------------------------------

LOG = logging.getLogger('web')

LANGUAGE = 'en-US'
ENCODING = 'utf-8'

UTF8 = codecs.getdecoder('utf-8')
WINDOWS = codecs.getdecoder('windows-1252')
LATIN1 = codecs.getdecoder('latin-1')


# ----- Web -------------------------------------------------------------------

class Web(object):
    session = None
    expiration = 180

    @classmethod
    def make_throttle_hook(cls, timeout=1.0):
        """
        Returns a response hook function which sleeps for `timeout` seconds if
        response is not cached
        """

        def hook(response, **kwargs):
            if not getattr(response, 'from_cache', False):
                LOG.debug('Cache throttling; %g seconds' % timeout)
                time.sleep(timeout)
            return response

        return hook

    @classmethod
    def load(cls, url, cached=True):
        if not cls.session:
            cls.session = requests_cache.CachedSession('cache', expire_after=timedelta(days=cls.expiration))
            cls.session.hooks = {'response': cls.make_throttle_hook(0.1)}

        if cached:
            response = cls.session.get(url)
        else:
            response = requests.get(url)

        return response

    @classmethod
    def parse(cls, html):
        if len(html):
            try:
                document = BeautifulSoup(html, 'lxml')
                return document
            except Exception as e:
                LOG.warn('Parsing exception: %s' % e)
                return ''
        else:
            return ''

    @classmethod
    def decode(cls, data):
        raw = ''
        try:
            raw = UTF8(data)[0]
            LOG.debug('Parsed %d bytes as %d UTF-8 characters' % (len(data), len(raw)))
        except UnicodeDecodeError:
            try:
                raw = WINDOWS(data)[0]
                LOG.debug('Parsed %d bytes as %d Windows-1252 characters' % (len(data), len(raw)))
            except UnicodeDecodeError:
                try:
                    raw = LATIN1(data)[0]
                    LOG.debug('Parsed %d bytes as %d Latin-1 characters' % (len(data), len(raw)))
                except UnicodeDecodeError as e:
                    LOG.warn('Parsing exception: %s' % e)
        return raw

    @classmethod
    def fetch(cls, url):
        response = cls.load(url)
        if response.status_code == 404:
            return None
        return cls.parse(response.content)


if __name__ == '__main__':
    parsed = Web.fetch('http://www.boardgamers.org/')
    print(parsed.prettify())
