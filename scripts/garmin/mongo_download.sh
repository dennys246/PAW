
var subjects = $1
var output_path = $2

# Open a connection to notre dames servers
ssh dschaed3@tesserae-wearable.crc.nd.edu

mongoexport -

"""
# Load the mongodb environment
mongo -u denny

# Connect to the garmin database
use garmin

# Iterate through subjects of interest
for (id in subjects){
	var oauth_token = db.tokens.find({'email':id}).toArray()[0]['oauth_token']

	var json = {
		'dailies' : db.dailies.find({'userAccessToken':oauth_token}).toArray(),
		'activities' : db.activies.find({'userAccessToken':oauth_token}).toArray(),
		'epochs' : db.epochs.find({'userAccessToken':oauth_token}).toArray(),
		'activity_details' : db.activityDetails.find({'userAccessToken':oauth_token}).toArray()
	}

	payload[id] = json
}"""
