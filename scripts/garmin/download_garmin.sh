#!/bin/bash

# Open a connection to notre dames servers and run export tokens from mongodb to generate subject pool text file
cat export_tokens.sh | ssh dschaed3@tesserae-wearable.crc.nd.edu 

# Sync Notre Dame server with local drive
rsync -r dschaed3@tesserae-wearable.crc.nd.edu:PAW/ ../../sourcedata/garmin/


