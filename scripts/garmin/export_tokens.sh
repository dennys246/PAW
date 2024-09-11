#!/bin/bash

mongoexport -d garmin -c tokens -o "PAW/tokens.csv" -u=denny -p='?Max!Autumn?!250' --authenticationDatabase=admin

python3 scripts/garmin/process_tokens.py

bash scripts/garmin/export_subjects.sh

exit
