# Foodsoft Importer for Artikles

## Feichtinger
Works right now only for Feichtiner..

_Special Conversions_

* loose articels with units in 'kg' are changed to 500g and prices gets adapted
* all `zukauf` articles are skipped
* also Bestellnummer over _100000_ is skipped, mostly meat and stuff

# Usage
## Import from Feichtinger to Foodsoft
* Download Feichtinger xls and give as param -i 
* Article list for foodsoft will be written to a file specified with param -o
* upload to file to foodsoft
## From foodsoft to Feichtinger
* Download Articles from foodsoft and provide path to param -b
* also provide the feichtiner xsl 
* output is a new xls copied from feichtinger xls with the values filled out, filename starts with KWxx_
* Original xls will not be overwritten or modified

## csv formats
foodsoft download:
````
Bestellte Gebinde;Bestellnummer;Name;Einheit;GebGr;Nettopreis;Total price
`````
feichtinger: 

# Params
````
usage: feichtinger-import-export.py [-h] [-i IN_FILE] [-o [OUT_FILE]] [-b [BESTELLUNG]] [-w WEEK]

Convert Bestellungen to Foodsoft

optional arguments:
  -h, --help            show this help message and exit
  -i IN_FILE, --in_file IN_FILE
                        Input file vom Lieferanten
  -o [OUT_FILE], --out_file [OUT_FILE]
                        Das csv zum Upload in die foodsoft
  -b [BESTELLUNG], --bestellung [BESTELLUNG]
                        Das csv mit den Bestellungen
  -w WEEK, --week WEEK  Defaults to the current week number
`````
