# Calendar makefile

SITE=http://boardgamers.org/downloads/
SCHEDULE=schedule
YEAR=2012
EXT=.xls

SPREADSHEET=$(SCHEDULE)$(YEAR)$(EXT)
OLD_SPREADSHEET=$(SCHEDULE)$(YEAR).old
NEW_SPREADSHEET=$(SCHEDULE)$(YEAR).new

BUILD=build

all:	build

.PHONY:	build clean compare fetch prerequisites push backup

build:	clean
	python WBC.py -i $(SCHEDULE)$(YEAR)$(EXT) -o $(BUILD)
	tar cf - Makefile WBC.py wbc-*-codes.csv wbc-template.html lgpl.txt | gzip > $(BUILD)/wbc-calendars.tar.gz

push:	build
	rsync -v -rlD --no-times --delete $(BUILD)/ trader.name:/data/web/trader/wbc/$(YEAR)/

clean:
	rm -rf $(BUILD)
	mkdir $(BUILD)

backup:
	rm -rf save
	mkdir save
	rsync -av $(BUILD)/ save/

fetch:
	wget -nv -O $(NEW_SPREADSHEET) $(SITE)$(SPREADSHEET)
	@if cmp -s $(NEW_SPREADSHEET) $(SPREADSHEET) ; then \
	  rm $(NEW_SPREADSHEET) ; \
	  echo "No changes to $(SPREADSHEET)" ; \
	else \
	  [ -f $(SPREADSHEET) ] && mv $(SPREADSHEET) $(OLD_SPREADSHEET) ; \
	  mv $(NEW_SPREADSHEET) $(SPREADSHEET) ; \
	  echo "$(SPREADSHEET) updated" ; \
	fi

prerequisites:
	sudo apt-get install python-pip
	sudo pip install --upgrade pytz
	sudo pip install --upgrade BeautifulSoup
	sudo pip install --upgrade icalendar
	sudo pip install --upgrade xlrd
