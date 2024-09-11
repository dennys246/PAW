import os, subprocess, csv, glob, pyminizip, zipfile, datetime, openpyxl, random
import pandas as pd
from datetime import datetime

class Processor:
	def __init__(self, debug = True, automatic = False):
		self.debug = debug
		self.automatic = automatic
		self.working_dir = '../../sourcedata/'
		return

	def __repr__(self):
		return

	def merge(self):
		self.garmin = Garmin()
		self.garmin.process()
		paw = self.garmin.data
		del self.garmin

		self.ema = EMA()
		paw['ema'] = self.ema.load()
		
		column = 2
		row = 0
		headers = {}
		row_dict = {}
		contents = []
		
		random_samples = []
		current_subject = None
		subject_start = None

		for source, subjects in paw.items(): # Iterate through all subjects
			for subject, datatypes in subjects.items(): # Iterate through all datatypes available
				for datatype, data in datatypes.items(): # Iterate through data available
					if datatype not in headers.keys():
						headers[datatype] = column
						column += 1
					if isinstance(data, dict) == False:
						continue
					for timestamp, datum in data.items():
						row_identifier = str(timestamp) + '-' + str(subject) 
						if row_identifier not in row_dict.keys():
							contents.append([])
							row_dict[row_identifier] = row
							row += 1
						datum_row = row_dict[row_identifier]

						row_diff = len(headers.keys()) - len(contents[datum_row]) + 2
						if row_diff > 0:
							contents[datum_row] += [None]*row_diff
						contents[datum_row][0] = timestamp
						contents[datum_row][1] = subject
						contents[datum_row][headers[datatype]] = datum

						# Randomly sample data
						if random.randint(0, 100) > 96:
							sample_start = (datum_row - random.randint(0,datum_row%100))
							random_samples += contents[sample_start:datum_row] 
				

		column_count = len(headers)
		for row_ind, row in enumerate(contents):
			row_diff = column_count - len(contents[row_dict[row_identifier]]) + 2
			if row_diff > 0:
				contents[row_ind] += [None]*row_diff

		# Save the data to the mastersheet
		with open(f'{self.working_dir}PAWMasterSheet.csv', 'w') as file:
			writer = csv.writer(file)
			writer.writerow(['Timestamp', 'Subject'] + list(headers.keys()))
			writer.writerows(contents)

		# Save the random sample
		with open(f'{self.working_dir}PAWSample.csv', 'w') as file:
			writer = csv.writer(file)
			writer.writerow(['Timestamp', 'Subject'] + list(headers.keys()))
			writer.writerows(random_samples)
			
		# Sort
		df = pd.read_csv(f'{self.working_dir}PAWMasterSheet.csv')
		df = df.sort_values(["Subject", "Timestamp"], )
		df.to_csv(f'{self.working_dir}PAWMasterSheet.csv')


	def zip(self, filename, new_filename = None, destination = None, compression_level = 6):
		if new_filename is None:
			new_filename = filename
		if destination is None:
			destination = self.working_dir
		try:
			with zipfile.ZipFile(new_filename.split('.')[0] + '.zip', 'w', compression = zipfile.ZIP_DEFLATED, compressionlevel = compression_level ) as zip_file:
				zip_file.write(destination + filename, arcname = new_filename)
			if self.automatic is True:
				return True
		except:
			if self.automatic is False:
				print(f'Script failed to zip the {filename} file')
			else:
				return False

	def unzip(self, filename, destination = None):
		if destination is None:
			destination = self.working_dir
		try:
			with zipfile.ZipFile(filename) as zip_file:
				zip_file.extractall(destination)
			if self.automatic is True:
				return True
		except:
			if self.automatic is False:
				print(f'Script failed to unzip the {filename} file')
			else:
				return False

	def encrypt(self, filename, parent, new_filename = None, destination = None, compression_level = 9):
		password = getpass(f"Encrypting file {filename}...\nPassword:")
		if new_filename is None:
			new_filename = filename.split('.')[0] + '.zip'
		if destination is None:
			destination = self.working_dir
		try:
			pyminizip.compress(parent + filename, None, destination + new_filename, password, compression_level)
			if self.automatic is True:
				return True
		except:
			if self.automatic is False:
				print(f'Script failed to encrypt the {filename} file')
			else:
				return False

	def decrypt(self, filename, compression):
		password = getpass(f"Decrypting file {filename}...\nPassword:")
		try:
			pyminizip.decompress()
			if self.automatic is True:
				return True
		except:
			if self.automatic is False:
				print(f'Script failed to decrypt the {filename} file')
			else:
				return False

	def bash(self, command):
		process = subprocess.Popen(command.split(), stdout =subprocess.PIPE)
		return process.communicate()

class Garmin(Processor):

	def __init__(self, debug = True, automatic = False):
		self.debug = debug
		self.automatic = automatic
		self.working_dir = '../../sourcedata/garmin/'
		return

	def transfer(self):
		# Need to establish ssh-keygen with petalibrary first before functional
		
		# Initiate vpn connection with ND
		self.bash('bash ../garmin/download_garmin.sh')

		# Initiate vpn connection with CU
		# rsync with peta library
		return
	
	def process(self):
		self.data = {}

		self.data['activities'] = self.process_activities()
		self.data['dailies'] = self.process_dailies()
		self.data['pulseox'] = self.process_pulseox()
		self.data['sleep'] = self.process_sleep()
		return
	
	def process_dailies(self):
		dailies = {}
		filepaths = glob.glob(f'{self.working_dir}*/*dailies.csv')
		for path in filepaths:
			data = {}
			filename = os.path.basename(path)
			subject = filename.split('-')[0]
			with open(path, 'r') as file:
				reader = csv.reader(file)
				for row in reader:

					start_second = None
					hr_found = False
					for item in row:
						
						if item[:18] == 'startTimeInSeconds':
							start_second = int(item.split(':')[1])

						if item[:15] == 'timeOffsetHeart':
							hr_found = True
							item = item.split('{')[1]
							datum = item.split(':') # Split entry into datum
							if '"' in datum[0]: # Remove "" from first entry
								datum[0] = datum[0].split('"')[1]
							if datum[0] == '}': # If no data present
								break # Out of for loop
						elif hr_found == True:
							datum = item.split(':')
							datum[1] = datum[1].split('}')[0]
						else: # If we haven't found HR yet
							continue

						data[int(datum[0]) + start_second] = int(datum[1]) # Add data to subject dictionary

						if item[-1] == '}':
							break
			dailies[subject] = {}
			dailies[subject]['Heart Rate'] = data
		return dailies
	
	def process_activities(self):
		# Data entry {second : [heart_rate, ]}
		activities = {}
		filepaths = glob.glob(f'{self.working_dir}*/*activities.csv')
		for path in filepaths:
			data = {}
			filename = os.path.basename(path)
			subject = filename.split('-')[0]
			with open(path, 'r') as file:
				reader = csv.reader(file)
				for activity in reader:
					
					datum = []
					headers = []
					hr_found = False
					for item in activity:

						if item[:18] == 'startTimeInSeconds':
							sample_time = int(item.split(':')[1])

						if item[:9] == 'samples:[': # If beginning of sample
							hr_found = True
							item = item.split('[')[1]

						if item == ']': # If end of sample
							continue

						if hr_found: # If we are looking at samples
							temp = item

							if temp[0] == '{': # Remove curly brackets
								temp = temp[1:]
							if temp[-1] == '}':
								temp = temp[:-1] 

							split = temp.split(':')
							if split[0][0] == '"':
								split[0] = split[0][1:-1] # Remove comma's
							headers.append(split[0])
							datum.append(split[1])

						if item[-1] == '}' and hr_found: # Collect the sample and add to data
							for data_ind, header in enumerate(headers):
								if header not in data.keys():
									data[header] = {}
								data[header][sample_time] = datum[data_ind] # Add data to subject dictionary
							headers = []
							datum = []
			activities[subject] = data
		return activities
	
	def process_pulseox(self):
		pulseox = {}
		filepaths = glob.glob(f'{self.working_dir}*/*pulseox.csv')
		for path in filepaths:
			data = {}
			filename = os.path.basename(path)
			subject = filename.split('-')[0]
			with open(path, 'r') as file:
				reader = csv.reader(file)
				for row in reader:

					start_second = None
					hr_found = False
					for item in row:
						
						if item[:18] == 'startTimeInSeconds':
							start_second = int(item.split(':')[1])

						if item[:20] == 'timeOffsetSpo2Values':
							hr_found = True
							item = item.split('{')[1]
							datum = item.split(':') # Split entry into datum
							if '"' in datum[0]: # Remove "" from first entry
								datum[0] = datum[0][1:-1]
							if datum[0] == '}': # If no data present
								break # Out of for loop
						elif hr_found == True:
							datum = item.split(':')
						else: # If we haven't found HR yet
							continue

						datum[1] = datum[1].split('}')[0] # Remove any curly brackets

						data[int(datum[0]) + start_second] = int(datum[1]) # Add data to subject dictionary

						if item[-1] == '}':
							break
							
			pulseox[subject] = {}
			pulseox[subject]['Pulseox'] = data
		return pulseox

	def process_sleep(self):
		sleep = {}
		filepaths = glob.glob(f'{self.working_dir}*/*sleep.csv')
		for path in filepaths:
			data = {'sleep_spo2': {}, 'awake': {}, 'light_sleep': {}, 'deep_sleep': {}}
			filename = os.path.basename(path)
			subject = filename.split('-')[0]
			with open(path, 'r') as file:
				reader = csv.reader(file)
				for row in reader:

					start_second = None
					spo2_found = False
					for item in row:
						if item[:15] == 'userAccessToken':
							break
						
						if item[:18] == 'startTimeInSeconds':
							item = ''.join(item.split('}'))
							item = ''.join(item.split(']'))
							start_second = int(item.split(':')[1])

						if item[:19] == 'timeOffsetSleepSpo2':
							spo2_found = True
							item = item.split('{')[1]
							datum = item.split(':') # Split entry into datum
							if '"' in datum[0]: # Remove "" from first entry
								datum[0] = datum[0][1:-1]
							if datum[0] == '}': # If no data present
								break # Out of for loop
						elif spo2_found == True:
							datum = item.split(':')
						else: # If we haven't found HR yet
							continue
						
						datum[1] = datum[1].split('}')[0] # Remove any curly brackets
						data['sleep_spo2'][int(datum[0]) + start_second] = int(datum[1]) # Add data to subject dictionary

					sleep_found = False
					for item in row: # Look for 
						if item[:6] == 'userId':
							break
						
						if item[:18] == 'startTimeInSeconds':
							item = ''.join(item.split('}'))
							item = ''.join(item.split(']'))
							start_second = int(item.split(':')[1])

						if item[:14] == 'sleepLevelsMap':
							sleep_found = True
							item = '{'.join(item.split('{')[1:]) # Remove first portion
							datum = item.split(':') # Split entry into datum
							if len(datum) > 2:
								condition = datum[0]
								datum = datum[1:]
							if '"' in datum[0]: # Remove "" from first entry
								datum[0] = datum[0][1:-1]
							if datum[0] == '}': # If no data present
								break # Out of for loop
						elif sleep_found == True:
							datum = item.split(':')
							if len(datum) > 2:
								condition = datum[0]
								datum = datum[1:]
						else: # If we haven't found HR yet
							continue
						
						if len(datum[0].split('{')) > 1:
							datum[0] = ''.join(datum[0].split('{')[1:])

						datum[1] = datum[1].split('}')[0] # Remove any curly bracket

						if datum[0][:4] == '"end':
							datum[0] = 'end'
						elif datum[0][:5] == 'start':
							datum[0] = 'start'
						else:
							continue

						if condition not in data.keys():
							data[condition] = {}

						print(datum)
						
						data[condition][int(datum[1]) + start_second] = datum[0] # Add data to subject dictionary
											
			sleep[subject] = {}
			sleep[subject] = data
		return sleep
	
class EMA(Processor):

	def __init__(self, debug = True, automatic = False):
		self.debug = debug
		self.automatic = automatic
		self.working_dir = '../../sourcedata/ema/'
		return

	def process(self):
		return

	def process_data(self, subject_pool):
		'''
		# Create a master excel to add all data onto

		dir = os.listdir(self.working_dir)
		for file in dir:
			if file[-4:] == '.zip':
				self.unzip(file)

				# Merge with master excel

		# Figure out which participant is which
		'''
		ema = pd.read_csv(f'{self.working_dir}\EMA.csv')
		ema = ema.sort_values(by = ['Subject ID', 'Start Date'])# Sort excel based on subject, then date
		ema.to_csv('EMA.csv')

		ema = []
		with open('EMA.csv', newline = '\n') as file:
			csvreader = csv.reader(file, delimiter = ',')
			for row in file:
				ema.append(row.split(',')[1:-1])

		subject = subject_pool[0]
		for ind, datum in enumerate(ema[1:]):
			recorded_time = datetime.strptime(datum[0], '%m/%d/%Y %H:%M%p').replace(tzinfo = None)
			if subject.SID != datum[6]: # Update subject if needed
				for subject in subject_pool: # Look for new subject info
					if subject.SID == int(datum[6]): # If found
						break # Break out of loop

			if subject.start_date:
				time_delta = recorded_time - subject.start_date
				ema[ind + 1][1] = time_delta.days  # Calculate day of study
				ema[ind + 1][2] = time_delta.seconds # Calculate second of study

		with open('..\sourcedata\ema\EMA-new.csv', 'w', newline = '\n') as file:
			writer = csv.writer(file)
			for datum in ema:
				writer.writerow(datum)

		self.data = ema

		return

	def load(self, filename = 'EMA.csv'):
		file = open(f'{self.working_dir}{filename}')
		reader = csv.reader(file)
		headers = next(reader)[2:]
		data = {}
		epoch_time = datetime(1970, 1, 1)
		for row in reader:
			if row[0][-7:-5] == '00':
				row[0] = '12:'.join(row[0].split('00:'))
			date = datetime.strptime(row[0], '%m/%d/%Y %I:%M%p')
			second_start = (date - epoch_time).total_seconds()
			if row[1] not in data.keys():
				data[row[1]] = {header : {} for header in headers}
			for ind, header in enumerate(headers):
				data[row[1]][header][second_start] = row[ind + 2]
		return data


class Earable(Processor):

	def __init__(self, debug = True, automatic = False):
		self.debug = debug
		self.automatic = automatic
		self.working_dir = '/earable/PAW/sourcedata/earable/'
		return

	def __repr__(self):
		return
