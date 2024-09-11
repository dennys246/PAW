import csv

subject_tokens = {}

with open('PAW/tokens.csv', 'r') as file:# Find token
    csvreader = csv.reader(file)
    for row in csvreader:
        subject_token = None
        for item in row: # Iterate through row
            if item[0:16] == 'email:"ucb-pain-': # If subject found
                subject_token = row # Declare the row as a subject token
                subject = item[16:-2] # Grab subject ID
                break # Exit for loop

        # If a subject token found
        if subject_token:
            for item in subject_token: # Iterate through token
                if item[0:12] == 'oauth_token:': # Find oauth token
                    oauth_token = item[13:-1] # Declare token
                    break # Exit for loop
            subject_tokens[subject] = oauth_token

with open('PAW/subject_pool.txt', 'w') as file:
    for subject, oauth_token in subject_tokens.items():
        file.write(f'{subject}/{oauth_token}\n')


