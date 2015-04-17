# Calendar makefile

SITE=http://boardgamers.org/downloads/
SCHEDULE=schedule
YEAR=2014
EXT=.xlsx
REMOTE=trader.name:/data/web/trader/wbc

SPREADSHEET=$(SCHEDULE)$(YEAR)$(EXT)
OLD_SPREADSHEET=$(SCHEDULE)$(YEAR).old
NEW_SPREADSHEET=$(SCHEDULE)$(YEAR).new

BUILD=site

all:	build

.PHONY:	build clean compare fetch prerequisites push backup

build:	
	python WBC.py -i $(SCHEDULE)$(YEAR)$(EXT) -o $(BUILD) -y $(YEAR)

dryrun:
	python WBC.py -i $(SCHEDULE)$(YEAR)$(EXT) -o $(BUILD) -n

pull:
	rm -rf live
	mkdir live
	rsync -v -rclD --delete $(REMOTE)/$(YEAR)/ live/

publish:
	rsync -v -rclD --delete $(BUILD)/ $(REMOTE)/$(YEAR)/

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
	sudo pip install --upgrade icalendar
	sudo pip install --upgrade beautifulsoup4
	sudo pip install --upgrade xlrd
