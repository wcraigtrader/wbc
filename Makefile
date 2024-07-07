# Calendar makefile

SITE=http://boardgamers.org/downloads/
SCHEDULE=schedule
YEAR=2024
EXT=.xlsx
REMOTE=trader.name:/data/web/trader/wbc
# REMOTE=wbc.trader.name:/var/www/wbc/schedule

# SPREADSHEET="2023 WBC Schedule - App vFinal2.xlsx"
# SPREADSHEET="2022 WBC Schedule - App v2.0 Final.xlsx"
SPREADSHEET="$(SCHEDULE)-$(YEAR)$(EXT)"
OLD_SPREADSHEET=$(SCHEDULE)$(YEAR).old
NEW_SPREADSHEET=$(SCHEDULE)$(YEAR).new

BUILD=build
CACHE=cache

all:	build

.PHONY:	build clean compare fetch prerequisites push backup

build:	
	python WBC.py -t new -i $(SPREADSHEET) -o $(BUILD)

dryrun:
	python WBC.py -d -t new -i $(SPREADSHEET) -o $(BUILD) -n

pull:
	rm -rf live
	mkdir live
	rsync -v -rclD --delete $(REMOTE)/$(YEAR)/ live/

publish:
	rsync -rclD --delete $(BUILD)/ $(REMOTE)/$(YEAR)/

clean:
	rm -rf $(BUILD)
	mkdir $(BUILD)

very-clean: clean
	rm -rf $(CACHE)
	mkdir $(CACHE)

backup:
	rm -rf save
	mkdir save
	rsync -a $(BUILD)/ save/

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
	conda install -y \
		beautifulsoup4 \
		openpyxl \
		lxml \
		requests_cache \
		icalendar

pip-prerequisites:
	sudo -H pip install --upgrade icalendar
	sudo -H pip install --upgrade beautifulsoup4
	sudo -H pip install --upgrade xlrd
	sudo -H pip install --upgrade lxml
	sudo -H pip install --upgrade requests_cache
