# Calendar makefile

# SRCS=WBCScheduleSpreadsheet
SRCS=schedule
YEAR=2011

all:	build

.PHONY:	build clean compare fetch prerequisites push

build:	clean
	python WBC.py -i $(SRCS)$(YEAR).xls
	tar cf - Makefile WBC.py wbc-*-codes.csv wbc-template.html lgpl.txt | gzip > test/wbc-calendars.tar.gz

compare:
	@$(SHELL) compare-live-test

push:
	rsync -v -rltD --delete test/ trader.name:/data/web/trader/wbc/$(YEAR)/
	# rsync -v -rltD --delete test/ live/$(YEAR)/

clean:
	rm -rf test
	mkdir test

test:	build compare

fetch:
	wget -nv -O $(SRCS)$(YEAR).new http://boardgamers.org/downloads/$(SRCS)$(YEAR).xls
	@if cmp -s $(SRCS)$(YEAR).new $(SRCS)$(YEAR).xls ; then \
	rm $(SRCS)$(YEAR).new ; \
	echo "No changes to $(SRCS)$(YEAR).xls" ; \
	else \
	mv $(SRCS)$(YEAR).xls $(SRCS)$(YEAR).old ; mv $(SRCS)$(YEAR).new $(SRCS)$(YEAR).xls ; \
	echo "$(SRCS)$(YEAR).xls updated" ; \
	fi

prerequisites:
	sudo apt-get install python-pip
	sudo pip install --upgrade pytz
	sudo pip install --upgrade BeautifulSoup
	sudo pip install --upgrade icalendar
	sudo pip install --upgrade xlrd
