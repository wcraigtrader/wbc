# WBC: Generate iCal calendars from the WBC Schedule spreadsheet

## PURPOSE:

The World Boardgaming Championships (http://boardgamers.org/#wbc) is a huge
annual boardgame tourney, with more than 100 tournaments running over the 
course of 9 days.  While the WBC produces their event schedule in many forms,
they don't produce iCal-style web calendars for the tournaments.  This program
is intended to fill the gap.

## PREREQUISITES:

* BeautifulSoup4
* icalendar
* requests
* requests_cache
* openpyxl

## USAGE:

Install the prerequisites and then run WBC from the command line:

	mkdir build
	python WBC.py

This will download the available calendar references, compare them,
and generate the appropriate iCal calendars in the build directory.

## HISTORY

From the calendar perspective, a WBC tournament is comprised of one or more optional demo events, possibly a mulligan event, a combination of heats and elimination rounds, leading up to a semi-final and final.

When I started attending WBC in 2007, the schedule was available in several forms, but none of them contained all of the schedule data in one place. There was an HTML all-in-one calendar that only listed event start times, nominal lengths, and a color code for the type of event, web pages for each scheduled tournament that included all of the information about each tournament, inconsistently coded with icons and text, and there were a couple of PDF documents that were intended to be printed, not parsed. The original data for the schedule was encoded in a Word document and maintained by hand, but not distributed.

I wanted to be able to load individual tournaments into my online calendar to help in planning which events to attend, so in 2009 I created this tool. I started with scraping the web pages for the all-in-one tournament guide and the individual tournaments and generating iCal (webcal) files for each tournament that concerned me. 

Starting in 2012, the tournament director started releasing the schedule data in an XLS spreadsheet, so I moved to interpreting the spreadsheet data.

In 2013, I started doing quality control for WBC, comparing all of the online schedules against the spreadsheet data. In addition, the spreadsheet moved from XLS to XSLX format, which required some rework.

In 2016, scheduling apps for WBC started to be released, so I took the cleansed event data for my webcals and started generating CSV and JSON data that could be used by the app developers. At the same time, the convention moved to a new site, with different room names. At the same time, the all-in-one calendar was phased out.

In 2018, the schedule spreadsheet moved to a one-line-per-event format (a great improvement) and the website arrangement changed significantly.

WBC wasn't held in 2020 or 2021, due to the Covid-19 pandemic.

In 2022, the spreadsheet layout changed in minor ways.

## NOTES:

When I started this project, Python 2.7 was the supported version of Python, and Python 3 was pretty buggy. A couple years later I attempted a conversion to Python 3 which aborted due to compatibility issues and time constraints. In 2022, the jump to Python 3 became a necessity, since Python 2.7 was EOL.

## LICENSE:

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published
by the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
