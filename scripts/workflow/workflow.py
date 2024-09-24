import connector, os, os.path, csv, sys, time, math, psutil, random, signal, atexit, math
import numpy as np
from datetime import datetime
from datetime import timedelta
from pytz import timezone
from pywintypes import TimeType
from xlrd import *
import win32com.client

class protocol: # Class to conduct workflow automation for the Pain Assessment with Wearables study
	def __init__(self, debug = True, cwd = None):
		print('Initializing PAW protocol...')
		self.debug = debug
		self.automatic = True

		if cwd != None:
			self.paw_cwd = cwd
		else:
			self.paw_cwd = "C:\GitSpot\PAW"

		print("Initializing PAW connector...")
		self.sys = connector.System()
		self.backlog = []

		self.today = datetime.today()
		self.today = self.today.replace(hour = 0)
		self.today = self.today.replace(tzinfo = timezone("US/Mountain"))

		self.participant_file = f'{self.paw_cwd}\PAW_Participants.xlsx'

		self.log_file = open(f'{self.paw_cwd}\logs\paw_{self.today.month}-{self.today.day}-{self.today.year}.log', 'w')
		self.log_file.write('- PAW Workflow Logs - ')

		print("Restarting Excel...")
		for process in psutil.process_iter():# Kill Microsoft excel process
			if 'EXCEL.EXE' in process.name():
				if self.automatic is False:
					answer = input('Excel seems to be open right now however to run this script excel must be closed, would you like to continue? (y/n)\n')
					if answer == 'y':
						os.kill(process.pid, signal.SIGTERM)
					else:
						return 'Script canceled...'
				else:
					os.kill(process.pid, signal.SIGTERM)

		print("Initialize protocol connection...")
		self.connect()
		atexit.register(self.handle_exit)

		self.subject_pool = []
		self.active_participants = []

		self.updates = []
		self.issues = []

		self.screening = []
		self.active = []

		self.positive_outcome = ['ok', 'a bit low', 'yes', 'accept']
		self.neutral_outcome = ['n/a', 'na', 'pause', 'wait']
		self.negative_outcome = ['no','pain too low', 'pain too low and infrequent', 'No to wearables during week - but might consider Garmin only?', 'dsq']

		self.participant_gdrive_id = '1OHW-XMuvKVryfrEGE8hpnTlRbu2HdhsT'

		print("Downloading data...")
		self.download_data() #, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		self.load_data()
		self.orient()

		print(f"PAW automation class instantiated, data loaded and oriented...\nSubject Count: {len(self.subject_pool)}\n Debug: {self.debug}", file = self.log_file)

	def handle_exit(self):
		results = self.save_data()
		self.disconnect()

	def connect(self):
		self.xlApp = win32com.client.Dispatch("Excel.Application")
		self.xlApp.DisplayAlerts = False

		#self.GDrive = connector.DriveConnection(self.debug, self.automatic)\
		self.Dropbox = connector.DropBoxConnection(self.log_file, working_directory = self.paw_cwd)
		self.Garmin = connector.GarminConnection(None, self.debug, self.automatic, self.log_file, working_directory = self.paw_cwd)
		self.Gmail = connector.GmailConnection(self.debug, self.automatic, self.log_file, working_directory = self.paw_cwd)

		#self.Twilio = connector.TwilioConnection(self.debug, self.automatic)

	def disconnect(self):
		#self.GDrive.disconnect()
		self.Garmin.disconnect()
		self.Dropbox.disconnect()

		#self.Peta.disconnect()
		#self.Mongo.disconnect()

		#self.Twilio.disconnect()


	def run(self):
		self.screen()
		self.schedule()
		self.compliance()
		self.report()


	def orient(self):
		participant_ind = 3
		print("Orienting script...")
		while self.participants.Cells(participant_ind, 1).Value is not None:
			SID = self.participants.Cells(participant_ind, 1).Value
			if isinstance(SID, float) == True and SID != None: # If the individual has been given a subject ID
				SID = int(SID)
				if self.participants.Cells(participant_ind, 29).Value not in [None, '']:
					participant_ind += 1
					continue

				session_date = self.participants.Cells(participant_ind, 12).Value
				if session_date in [None, '']:
					participant_ind += 1
					continue
				else:
					if type(session_date) is TimeType: # If consent date available convert to datetime object
						session_date = datetime.fromtimestamp(timestamp = session_date.timestamp(), tz = None)
						if (self.today.replace(tzinfo=None) - session_date).days < 1:
							participant_ind += 1
							continue


				first_name = self.participants.Cells(participant_ind, 2).Value
				if first_name is not None:
					last_name = self.participants.Cells(participant_ind, 3).Value

					consented = self.participants.Cells(participant_ind, 7).Value # Grab date consented
					if consented is not None: # If consent date available convert to datetime object
						consented = datetime.fromtimestamp(timestamp = consented.timestamp(), tz = consented.tzinfo)

					start_date = self.participants.Cells(participant_ind, 19).Value
					if type(start_date) is TimeType: # If consent date available convert to datetime object
						start_date = datetime.fromtimestamp(timestamp = start_date.timestamp(), tz = None)
					else:
						start_date = None

					# Iterate through the eligibilty sheet till there are two consecutive missing IP addresses
					eligibility_ind = 5
					email, number, prefered_contact = None, None, None
					while self.eligibility.Cells(eligibility_ind, 8).Value is not None and self.eligibility.Cells(eligibility_ind + 1, 8) is not None:
						e_first_name = self.eligibility.Cells(eligibility_ind, 24).Value
						e_last_name = self.eligibility.Cells(eligibility_ind, 23).Value
						if e_first_name == first_name and e_last_name == last_name: # If the current entry is the consented individual
							break
						eligibility_ind += 1

					subject = participant(SID, first_name, last_name, email, number, prefered_contact, consented)
					subject.eligibility_row = eligibility_ind
					subject.participant_row = participant_ind
					subject.start_date = start_date
					self.subject_pool.append(subject)
			participant_ind += 1

	def screen(self):
		# iterate through eligibility and send eligility emails
		screening_ind = 5
		while self.screeners.Cells(screening_ind, 1).Value is not None:
			print(screening_ind, file = self.log_file)
			print(screening_ind)

			# Grab name and process
			first_name = str(self.screeners.Cells(screening_ind, 1).Value)
			if first_name == None: # If it's an empty response
				screening_ind += 1
				continue

			first_name = first_name.strip(' ')
			first_name = first_name[0].upper() + first_name[1:].lower()

			last_name = str(self.screeners.Cells(screening_ind, 2).Value).strip(' ')
			last_name = last_name[0].upper() + last_name[1:].lower()


			# Check email provided
			email = str(self.screeners.Cells(screening_ind, 3).Value)
			if email == None:
				print(f"{first_name} email is missing, moving to next participant...")
				screening_ind += 1
				continue
			email = email.strip(' ')


			# Grab script markers within file
			first_invitation_sent = self.screeners.Cells(screening_ind, 4).Value

			print(last_name)
			if first_invitation_sent != None:
				if isinstance(first_invitation_sent, str) and first_invitation_sent.lower().strip() in self.neutral_outcome + self.negative_outcome:
					screening_ind += 1
					continue

				if isinstance(first_invitation_sent, str):
					self.issues.append(f"Participant in row {screening_ind} of eligibility sheet does not have a date inputted as their Intial email sent (column B), second screen cannot be send")
					screening_ind += 1
					continue

				if first_invitation_sent in self.neutral_outcome:
					screening_ind += 1
					continue

				first_invitation_sent = datetime.fromtimestamp(timestamp = first_invitation_sent.timestamp(), tz = first_invitation_sent.tzinfo)
				intro_difference = round((self.today - first_invitation_sent).total_seconds() / (60*60*24), 0)

			# ----------------------------- Send Screening Form --------------------------------#
			if first_invitation_sent not in [None, 'pause', 'n/a', 'stop', 'withdrawn']:
				time_delta = round((self.today - first_invitation_sent).total_seconds() / (60*60*24), 0) # Find days between invitation being sent and today
				print(f"{screening_ind} - {time_delta}")
				# Find last screening sent
				screen_count = 0
				last_sent = self.screeners.Cells(screening_ind, 5).Value
				next_sent = self.screeners.Cells(screening_ind, 5).Value
				while next_sent not in ['', None, 'None']: # Find next survey to be sent
					last_sent = next_sent
					screen_count += 1
					if screen_count >= 6: # If we've reached our last screening
						break # Exit loop

					next_sent = self.screeners.Cells(screening_ind, 5 + screen_count).Value # Update next step

				screen_count += 1 # Increment screening count

				if screen_count > 6: # Go to next person if we've reached our max screening
					screening_ind += 1 #
					continue

				if isinstance(last_sent, datetime):
					last_sent = datetime.fromtimestamp(timestamp = last_sent.timestamp(), tz = last_sent.tzinfo)
					time_delta = round((self.today - last_sent).total_seconds() / (60*60*24), 0)
				elif isinstance(last_sent, str):
					self.issues.append(f'Subject in row {screening_ind} has a string instead of date in sheet')

				if time_delta >= 2 and time_delta <= 9:

					if time_delta == 2 and random.randint(0, 1) == 1:
						screening_ind += 1
						continue # Randomly skip this day

					self.screening.append(f'Participant in row {screening_ind} - Survey {screen_count}')

					subject = "Screening Form to Complete Today"
					body = f"Hello {first_name},<br><br>"
					body += "Here is your <a href='https://cuboulder.qualtrics.com/jfe/form/SV_etSNrS6k1zbfCCO'>screening form</a> to complete today, please let us know if you have any questions.<br><br>Sincerely,<br>PAW Research Team"

					self.screeners.Cells(screening_ind, 4 + screen_count).Value = self.today

					print(body, file = self.log_file)
					if self.debug is False:
						self.Gmail.send_email(email, subject, body, 'html')
					if screen_count == 6:
						self.updates.append(f'Participant in row {screening_ind} has been sent their final second screen')

			screening_ind += 1


	def schedule(self):
		for participant in self.subject_pool:
			ind = participant.participant_row

			session_date = self.participants.Cells(ind, 12).Value

			if isinstance(session_date, str) and session_date.lower().strip() in self.neutral_outcome:
				continue

			if session_date == None or session_date in self.negative_outcome:
				continue

			session_date = datetime.fromtimestamp(timestamp = session_date.timestamp(), tz = session_date.tzinfo)
			session_time = self.participants.Cells(ind, 13).Value
		return

	def compliance(self):
		print("Assessing compliance...")
		usage = np.zeros((len(self.subject_pool), 21, 6))
		complete = 0
		in_progress = 0

		today = datetime.today()
		total_days_assessed = 4
		last_comparison = today - timedelta(days = total_days_assessed)

		daily_hour_goal = 8
		complaince_threshold = 0.70 # Percent compliant needed to trigger a warning

		self.reports = []

		print(self.subject_pool)

		for subject in self.subject_pool:
			print(subject.SID)
			day_count = 0
			action = ""

			# --------- Evaluate ExpiWell --------- #
			survey_count = 0
			recent_survey_count = 0

			# --------- Evaluate Garmin -----------#
			garmin_total_hours = 0
			garmin__excess_hours = 0
			garmin_recent_hours = 0 # Hours of device usage within the last x days

			sample_duration = (2/60) # set too 2 minutes

			subject_metrics = self.Garmin.gather_metrics(subject.SID)

			if subject_metrics == False:
				self.issues.append(f"Failed to monitor garmin compliance of paw{subject.SID}")
				continue

			report = f'Compliance Report - {subject.SID}\n'
			for key, value in subject_metrics.items():
				report += f'	{key}: {round(float(value), 2)}\n'


			# ---------- Evaluate Earable ----------#
			earable_total_hours = 0
			earable_excess_hours = 0
			earable_recent_hours = 0 # Hours of device usage within the last x days

			self.reports.append(report)
		return

	def transfer_data(self, output_path = None, overwrite = False):
		self.mongo = connector.MongoConnection()
		self.mongo.connect()

		with open(f'{self.paw_cwd}/scripts/credentials/nd_creds.txt', 'r') as file:
			nd_creds = file.read().split('\n')

		username = nd_creds[0]
		password = nd_creds[1]

		self.mongo.bash(f"mongoexport -d garmin -c tokens -o 'tokens.csv' -u={username} -p='{password}' --authenticationDatabase=admin")

		if output_path == None:
			output_path = f"{self.paw_cwd}sourcedata\mongo"

		for subject in self.subject_pool:
			# Check if data has already been transfered

			# Create a folder for the subject in the Peta Library

			# Load Garmin data in from MongoDB
			transferables = {}

			subject_token = None
			with open(f'{self.paw_cwd}tokens.csv', 'r') as file:# Find token
				csvreader = csv.reader(file)
				for row in csvreader:
					for item in row:
						if item == 'email:"ucb-pain-' + subject.SID + '"':
							subject_token = row
							break

			if subject_token:
				print(f"{subject.SID} token found", file = self.log_file)
				for item in subject_token:
					if item[0:18] == 'oauth_token_secret': token_secret = item[19:-1]
					if item[0:11] == 'oauth_token': token = item[12:-1]
					query = "{'userAccessToken':" + str(token) + "}"
					for collection in ['activities']:
						self.mongo.bash("mongoexport -d garmin -c " + collection + " -o '../" + str(subject.SID) + "_" + collection + ".csv' -query=" + query + " -u=denny -p='2022Denny@ChangeMe' --authenticationDatabase=admin")
						if transferables[subject.SID] is None:
							transferables[subject.SID] = [f"{subject.SID}_{collection}.csv"]
						else:
							transferables[subject.SID].append(f"{subject.SID}_{collection}.csv")

		self.mongo.disconnect() # Disconnect from notre dame's servers
		for subject, files in transferables: #Transfer data off of notre dame servers
			if os.path.dir(f"mkdir {self.paw_cwd}/sourcedata/mongo/{subject}/"):
				self.mongo.bash(f"mkdir {self.paw_cwd}/sourcedata/mongo/{subject}/") # Create the subject directory
			for file in files:
				self.mongo.bash(f"scp dschaed3@tesserae-wearable.crc.nd.edu:/staging_area/{subject}_{collection}.csv {self.paw_cwd}/mongo/{subject}/{subject}_{collection}.csv") # Save as some file on the Peta Library

			# Load Earable data from GDrive
			# Save as some file on the Peta Library

			# Load ExpiWell data from zip file
			# Save as an individual excel sheet
		return

	def process_data(self):
		tokens = []
		ids = []
		with open(f"{self.paw_cwd}/sourcedata/mongo/tokens.csv") as file: # Load tokens
			csvreader = csv.reader(file)
			for line in csvreader:
				tokens.append(line)

		for subject in self.subject_pool:
			for ind, row in enumerate(tokens):
				for item in row:
					if item == 'email:"ucb-pain-' + subject.SID + '"':
						token = row
						break

				if token:
					ids.append(subject_token[0])


		self.dailies = []
		with open(f"{self.paw_cwd}/sourcedata/mongo/dailies.csv") as file:
			csvreader = csv.reader(file)
			for line in csvreader:
				if line[0] in ids:
					self.dailies.append(line)

	def report(self):
		subject = "PAW Workflow Report"
		print(self.screening, self.active_participants)
		if (len(self.issues) + len(self.updates) + len(self.active_participants) + len(self.screening)) + len(self.reports) == 0:
			print('No updates or issues found', file = self.log_file)
			return
		body = 'Hey Team,\n'
		if len(self.updates) > 0:
			body += "\nThe script reported back some updates on the study, here's what it found:\n"
			for update in self.updates: body += '	- ' + update + '\n'
		if len(self.screening) > 0:
			body+= "\nIt looks like we have some subjects being current screening, here's a list of the subjects currently being screening...\n"
			for screened in self.screening: body += '   - ' + screened + '\n'
		if len(self.active_participants) > 0:
			body += '\nIt looks like we have some active subjects collecting data...\n'
			for active in self.active_participants: body += '   - paw' + str(active) + '\n'
		if len(self.issues) > 0:
			body += "\nThe script picked up some issues with the current participants today, here's what it found:\n"
			for issue in self.issues: body += '	- ' + issue + '\n'
		if len(self.reports) > 0:
			body += "\nThe script generated some report metrics on actively running subjects. Here's the report(s):\n"
			for report in self.reports: body += report
		body += '\nBest,\nA Tiny Robot'

		print(body, file = self.log_file)
		if self.debug is False:
			self.Gmail.send_email('marta.ceko@colorado.edu;desc7849@colorado.edu;paola.badilla@colorado.edu', subject, body)

	def log(self, text):
		self.backlog.append(text)

	def load_data(self):
		self.participant_workbook = self.xlApp.Workbooks.Open(f'{self.paw_cwd}\PAW_Participants.xlsx', False, False, None, 'zagreb')
		self.participants = self.participant_workbook.Worksheets('Participants ')
		self.eligibility = self.participant_workbook.Worksheets('Eligibility Screeners')
		self.screeners = self.participant_workbook.Worksheets('Second Screeners')

	def download_data(self):
		self.Dropbox.download('/A_PAW/PAW_Participants.xlsx', f'{self.paw_cwd}\PAW_Participants.xlsx')

	# Old download command for downloading off Gdrive
	#def download_data(self, Id, filename, mime = None):
	#	stream = self.GDrive.load(Id, mime) # Download the io.BytesIO stream data from google
	#	with open(filename, 'wb') as file: # Save the IO stream data locally
	#		file.write(stream.getbuffer()) # Grab ByteIO stream buffer and write to file

	def save_data(self):
		#self.xlApp.DisplayAlerts = False
		self.participant_workbook.SaveAs(self.participant_file)
		time.sleep(3) # Wait for workbook to save before uploading to gdrive

		if self.debug == False:
			#self.GDrive.update(self.participant_gdrive_id, "../PAW_Participants.xlsx")
			self.Dropbox.upload(f'{self.paw_cwd}\PAW_Participants.xlsx', '/A_PAW/PAW_Participants.xlsx')


		print("Data saved locally & server connections cut...", file = self.log_file)

		self.log_file.close()

		for process in psutil.process_iter():# Kill Microsoft excel process
			if 'EXCEL.EXE' in process.name():
				os.kill(process.pid, signal.SIGTERM)



class participant:
	def __init__(self, SID, first_name, last_name, email, number, prefered_contact, consented):
		self.SID = SID
		self.status = None
		self.first_name = first_name
		self.last_name = last_name
		self.email = email
		self.number = number
		self.prefered_contact = prefered_contact

		self.consented = consented
		self.status = None
		self.start_date = None
		self.completion_date = None

		self.ema_identifier = None
		self.ema_subject_ids = []

		self.eligibility_row = None
		self.consent_row = None
		self.participant_row = None

debug = True
if len(sys.argv) > 1:
	if sys.argv[1] == 'False':
		debug = False


protocol = protocol(debug)
protocol.run()
