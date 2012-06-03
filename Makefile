# Calendar makefile

# SRCS=WBCScheduleSpreadsheet
SRCS=schedule
YEAR=2012
EXT=.xls

all:	build

.PHONY:	build clean compare fetch prerequisites push backup

build:	clean
	python WBC.py -i $(SRCS)$(YEAR)$(EXT)
	tar cf - Makefile WBC.py wbc-*-codes.csv wbc-template.html lgpl.txt | gzip > test/wbc-calendars.tar.gz

compare:
	@$(SHELL) compare-live-test

push:
	rsync -v -rlD --no-times --delete test/ trader.name:/data/web/trader/wbc/$(YEAR)/

clean:
	rm -rf test
	mkdir test

backup:
	rm -rf save
	mkdir save
	rsync -av test/ save/

test:	build compare

fetch:
	wget -nv -O $(SRCS)$(YEAR).new http://boardgamers.org/downloads/$(SRCS)$(YEAR)$(EXT)
	@if cmp -s $(SRCS)$(YEAR).new $(SRCS)$(YEAR)$(EXT) ; then \
	rm $(SRCS)$(YEAR).new ; \
	echo "No changes to $(SRCS)$(YEAR)$(EXT)" ; \
	else \
	mv $(SRCS)$(YEAR)$(EXT) $(SRCS)$(YEAR).old ; mv $(SRCS)$(YEAR).new $(SRCS)$(YEAR)$(EXT) ; \
	echo "$(SRCS)$(YEAR)$(EXT) updated" ; \
	fi

prerequisites:
	sudo apt-get install python-pip
	sudo pip install --upgrade pytz
	sudo pip install --upgrade BeautifulSoup
	sudo pip install --upgrade icalendar
	sudo pip install --upgrade xlrd
