#!/bin/bash

# Iterate through list generated from previous script
for line in $(cat PAW/subject_pool.txt) ; do
	# Split up line into token and subject
	IFS=/ read -r subject oauth_token <<< $line

	# Mongodb query to grab subjects token

	#query='"userAccessToken":"'$oauth_token'"'
	#query="'{$query}'"
	query='{"userAccessToken":"'$oauth_token'"}'

	# Check if subjects folder already exists on ND servers
	if [ ! -d "PAW/$subject/" ] ; then
		mkdir "PAW/$subject/"
	fi

	echo $query
	echo $subject
	# Export data from MongoDB to Notre Dame server
	mongoexport -d garmin -c pulseox -o "PAW/$subject/$subject-pulseox.csv" -u=denny -p='Denny@2022ChangeMe' -q=$query --authenticationDatabase=admin
	mongoexport -d garmin -c dailies -o "PAW/$subject/$subject-dailies.csv" -u=denny -p='Denny@2022ChangeMe' -q=$query --authenticationDatabase=admin
	mongoexport -d garmin -c activityDetails -o "PAW/$subject/$subject-activities.csv" -u=denny -p='Denny@2022ChangeMe' -q=$query --authenticationDatabase=admin
	mongoexport -d garmin -c sleeps -o "PAW/$subject/$subject-sleep.csv" -u=denny -p='Denny@2022ChangeMe' -q=$query --authenticationDatabase=admin
	mongoexport -d garmin -c stressDetails -o "PAW/$subject/$subject-stress.csv" -u=denny -p='Denny@2022ChangeMe' -q=$query --authenticationDatabase=admin
done
