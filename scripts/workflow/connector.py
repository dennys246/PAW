import io, os, urllib, pymongo, subprocess, csv, sys, time, math, dropbox
import shutil, psutil, signal, smtplib, atexit, pytz, math, pyminizip, zipfile
from garminconnect import Garmin, GarminConnectConnectionError, GarminConnectTooManyRequestsError, GarminConnectAuthenticationError
#from googleapiclient.discovery import build
#from googleapiclient.errors import HttpError
#from googleapiclient.http import MediaIoBaseDownload
#from googleapiclient.http import MediaFileUpload
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from sshtunnel import SSHTunnelForwarder
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from pytz import timezone
from twilio.rest import Client
from xlrd import *
from getpass import getpass
from paramiko import SSHClient
from scp import SCPClient
import win32com.client

class System:
	def __init__(self, debug = True, automatic = False, working_directory = None):
		self.debug = debug
		self.automatic = automatic
		self.today = datetime.today()

		self.working_directory = working_directory
		self.dataset_directory = self.working_directory

	def __repr__(self):
		return f"Date: {self.today}\nWorking Directory: {self.working_directory}\nDataset Directory: {self.dataset_directory}"

	def change(self, directory):
		os.chdir(directory)

	def list(self, directory):
		print("\n".join(os.listdir(directory)))

	def login(self, type, username = None):
		if username is None:
			username = input(f"Please enter your {type} username: ")
		print(f"Please provide the password too {username}'s {type} account... (Type 'new user' too change the user)")
		password = getpass()
		while "new user" in password.lower():
			username = input("What user would you like to log in as?")
			print(f"Please provide the password too {username}'s {type} account... (Type 'new user' too change the user)")
			password = getpass()
		return username, password

	def zip(self, filename, new_filename = None, destination = None, compression_level = 6):
		if new_filename is None:
			new_filename = filename
		if destination is None:
			destination = self.working_directory
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
			destination = self.working_directory
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
			destination = self.working_directory
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


class EmotivConnection(System):
	def __init__(self, debug = True, automatic = False, log_file = None, working_directory = None):
		atexit.register(self.disconnect)
		return

	def connect(self):
		# Grab credentials
		with open('../emotiv/.app_id', 'r') as file:
			self.client_id = file.read()[:-1]
		with open('../emotiv/.app_key', 'r') as file:
			self.client_secret = file.read()[:-1]
		self.client_secret = ''
		self.emotivpro_license = ''

		# Connect to cortex service
		self.cor = cortex.Cortex(self.client_id, self.client_secret)
		return

	def disconnect(self):
		return


class GarminConnection(System):
	def __init__(self, subject_id = None, debug = True, automatic = False, log_file = None, working_directory = None):
		self.working_directory = working_directory

		if subject_id is not None:
			self.connect(subject_id)
		else:
			self.api = None

		atexit.register(self.disconnect)

		self.activity_call = lambda date: self.api.get_activities(0, 30)
		self.general_call = lambda date: self.api.get_heart_rates(date.isoformat())

		if log_file != None:
			self.log_file = log_file

	def __del__(self):
		self.disconnect()

	def connect(self, subject_id):
		try:
			if str(subject_id)[:3] != 'paw':
				subject_id = f'paw{subject_id}'
			self.subject_id = subject_id
			self.api = Garmin(f"pawstudy.cu+{subject_id}@gmail.com", f"Zagreb@{subject_id}")
			self.api.login()
			return True
		except:
			return False

	def disconnect(self):
		if self.api:
			self.api.logout()
			del self.api

	def load(self, start_date = None, end_date = None, duration = 20):
		if end_date is None:
			end_date = datetime.today()
		if start_date is None:
			start_date = end_date - timedelta(days = duration)

		def daterange(start_date, end_date):
			for n in range(int((end_date - start_date).days)):
				yield start_date + timedelta(n)

		general_data = []
		activity_data = []
		for date in daterange(start_date, end_date):
			general_data.append(self.general_call(date))
			activity_data.append(self.activity_call(date))
		return general_data, activity_data

	def gather_metrics(self, subject_id = None, verbose = False):

		if subject_id != None:
			self.subject_id = subject_id
			status = self.connect(self.subject_id)
			if status == False:
				return status

		general_data, activity_data = self.load()

		metrics = {'Total Recorded Time (Hours)':0, 'Average Record Length (Hours)':0, 'Days Recorded':0, 'Recent Hours Recorded (3 Days)': 0}
		for day in general_data:
			# Skip if no recorded data
			if day['startTimestampGMT'] == None or day['heartRateValues'] == None:
				continue

			# Increment days recorded
			metrics['Days Recorded'] += 1

			# Calculate time recorded for day
			hr = day['heartRateValues']
			metrics['Total Recorded Time (Hours)'] += (hr[-1][0] - hr[0][0])/1000/60/60


		if metrics['Days Recorded'] > 0:
			metrics['Average Record Length (Hours)'] = metrics['Total Recorded Time (Hours)'] / metrics['Days Recorded']
		else:
			metrics['Average Record Length (Hours)'] = 0

		for day in general_data[-3:]:
			# Skip if no recorded data
			if day['startTimestampGMT'] == None or day['heartRateValues'] == None:
				continue

			# Calculate time recorded for day
			hr = day['heartRateValues']
			metrics['Recent Hours Recorded (3 Days)'] += (hr[-1][0] - hr[0][0])/1000/60/60

		if verbose == False:
			return metrics
		else:
			print(f'Compliance Report - {self.subject_id}')
			for key, value in metrics.items():
				print(f'{key}: {value}')




class MongoConnection(System):
	def __init__(self, debug = True, automatic = False, working_directory = None):
		self.debug = debug
		self.automatic = automatic

		self.working_directory = working_directory

		self.mongo_host = 'tesserae-wearable.crc.nd.edu'
		self.database_name = 'garmin'

		self.status = 'Not connected'


	def __repr__(self):
		return f"Mongo Host: {self.mongo_host}\nDatabase Name: {self.database_name}\nSSH User: {self.ssh_username}\nMongoDB User: {self.mongo_username}\nStatus: {self.status}"

	def __del__(self):
		if self.status != 'Disconnected':
			self.disconnect()

	def connect(self):
		#try:
		self.ssh_username, self.ssh_password = self.login('SSH', 'dschaed3')
		self.server = SSHTunnelForwarder(
			self.mongo_host,
			ssh_username = self.ssh_username,
			ssh_password = self.ssh_password,
			remote_bind_address = ('127.0.0.1', 8080)
		)
		self.server.start()
		print('Connection established too Notre Dame servers...')
		self.status = 'Connected'

		atexit.register(self.disconnect)

	def disconnect(self):
		if self.server:
			self.server.stop()
			self.server = None
			self.status = 'Disconnected'

class PetaConnection(System):
	def __init__(self, debug = True, automatic = False, working_directory = None):
		self.debug = debug
		self.automatic = automatic

		self.peta_ip = ''
		self.username = 'denny'
		print(f"Please provide the password too {self.username}'s account... (Type 'new user' too change the user)")
		self.password = getpass()
		if "new user" in self.password.lower():
			self.username = input("What user would you like to log in as?")
			print(f"Please provide the password too {self.username}'s account... (Type 'new user' too change the user)")
			self.password = getpass()

		self.status = self.connect()
		atexit.register(self.disconnect)

		self.working_directory = working_directory
		self.peta_directory = ''

		self.__repr__()

	def __repr__(self):
		return f"Peta Library IP: {self.peta_ip}\nPeta User: {self.username}\nPeta Directory: {self.peta_directory}\nWorking Directory: {self.working_directory}Status: {self.status}"

	def __del__(self):
		self.disconnect()

	def login(self, ip = None):
		if ip: self.peta_ip = ip
		self.username = input("Username:  ")
		self.password = getpass("Password:  ")
		self.connect()

	def connect(self, ip = None, username = None, password = None):
		if ip: self.peta_ip = ip
		if username: self.username = username
		if password: self.password = username

		self.ssh = paramiko.SSHClient()
		self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

		try:
			self.ssh.connect(self.peta_ip, username = self.username, password = self.password)
			if self.automatic is True:
				print("Connection successful...")
			else:
				return True
		except:
			if self.automatic is True:
				print("Connection too PetaLibrary failed...")
			else:
				return False

	def disconnect(self):
		if self.ssh:
			self.ssh.close()
			del self.ssh


	def sftp(self, filename, destination = None, new_filename = None):
		if destination is None: destination = self.peta_directory
		if new_filename is None: new_filename = filename
		try:
			self.sftp = self.ssh.open_sftp()
			try: # Check if the file already exists
				self.sftp.stat(destination + new_filename)
				if self.automatic is False:
					print("File already exists...")
				else:
					return False
			except: # Copy file over too destination
				self.sftp.put(filename, destination + new_filename)
				if self.automatic is False:
					print(f"File copied too {destination}")
				else:
					return True
		except paramiko.SSHException:
			print("Connection failed too PetaLibrary")

"""
	def connect(self, ip = None, username = None, password = None):
		if ip is not None: self.peta_ip = ip
		if username is not None: self.username = username
		if password is not None: self.password = password
		try:
			self.server = SSHTunnelForwarder(
				self.peta_ip,
				ssh_username = self.username,
				ssh_password = self.password,
				remote_bind_address = ('127.0.0.1', 8080))
			self.server.start()
			self.status = f'Successfully connected too {self.peta_ip} as {self.username}.'
		except:
			self.status = f'Failed too connect too Peta Library via {self.peta_ip} as {self.username}, try reconnecting using the call connection.connect(ip, username, password).'
		if self.automatic is False:
			self.__repr__()

"""
class DropBoxConnection(System):
	def __init__(self, debug = False, automatic = False, log_file = None, working_directory = None):
		self.debug = debug
		self.automatic = automatic

		self.working_directory = working_directory

		with open('../credentials/dropbox_creds.txt', 'r') as file:
			dropbox_creds = file.read().split('\n')

		self.app_key = dropbox_creds[0]
		self.app_secret = dropbox_creds[1]
		self.refresh_token = dropbox_creds[2]

		if log_file != None:
			self.log_file = log_file

		self.connect()
		atexit.register(self.disconnect)

	def connect(self):
		self.dpx = dropbox.Dropbox(
			#oauth2_access_token = self.access_token,
			app_key = self.app_key,
			app_secret = self.app_secret,
			oauth2_refresh_token = self.refresh_token,

		)

	def download(self, file, path):
		metadata, f = self.dpx.files_download(file)
		output = open(path, 'wb')
		output.write(f.content)
		output.close()

	def upload(self, file, path):
		with open(file, 'rb') as file:
			meta = self.dpx.files_upload(file.read(), path, mode = dropbox.files.WriteMode("overwrite"))
		return meta

	def disconnect(self):
		try:
			del self.dpx
		except:
			return

class DriveConnection(System):
	def __init__(self, debug = False, automatic = False, working_directory = None):
		self.debug = debug
		self.automatic = automatic

		self.working_directory = working_directory

		self.connect()
		atexit.register(self.disconnect)

	def connect(self):
		self.scopes = ['https://www.googleapis.com/auth/drive']
		self.creds = None
		if os.path.exists('../credentials/token.json'):
			self.creds = Credentials.from_authorized_user_file('../credentials/token.json', self.scopes)
		# If there are no (valid) credentials available, let the user log in.
		if not self.creds or not self.creds.valid:
		    if self.creds and self.creds.expired and self.creds.refresh_token:
		        self.creds.refresh(Request())
		    else:
		        flow = InstalledAppFlow.from_client_secrets_file('../credentials/credentials.json', self.scopes)
		        self.creds = flow.run_local_server(port=0)
		    # Save the credentials for the next run
		    with open('../credentials/token.json', 'w') as token:
		        token.write(self.creds.to_json())

		self.service = build('drive', 'v3', credentials = self.creds) # Call for an API instance

	def disconnect(self):
		if self.service != None:
			del self.service

	def load(self, Id, mime = None):
		try:
			if mime is None:
				request = self.service.files().get_media(fileId = Id)
			else:
				request = self.service.files().export_media(fileId = Id, mimeType = mime)
			file = io.BytesIO()
			downloader = MediaIoBaseDownload(file, request)
			done = False
			while done == False:
				status, done = downloader.next_chunk()

		except HttpError as error:
			if self.automatic is False:
				print(f'Script failed too download a file from google servers\n\nError: {error}')
			return False
		return file

	def save(self, filename, parent, destination_id, mimetype = 'application/vnd.ms-excel'):
		try:
			file_metadata = {'name': filename, 'parents': [destination_id], 'mimeType': mimetype}
			media = MediaFileUpload(parent + filename, mimetype= mimetype)

			if self.debug is False:
				file = self.service.files().create(body = file_metadata, media_body = media, fields = 'id').execute()
			return True

		except HttpError as error:
			if self.automatic is False:
				print(f'Script failed to upload {filename} too google servers\n\nError: {error}')
			return False

	def update(self, file_id, new_filename, mime_type = 'application/vnd.ms-excel'):
		try:
			# First retrieve the file from the API.
			file = self.service.files().get(fileId=file_id).execute()

			# File's new content.
			media_body = MediaFileUpload(new_filename, mimetype=mime_type)

			# Send the request to the API.
			if self.debug is False:
				updated_file = self.service.files().update(fileId = file_id, media_body = media_body).execute()
			if self.automatic is True:
				return True
		except HttpError as error:
			if self.automatic is False:
				print('An error occurred: %s' % error)
			else:
				return False

	def move(self, file_id, folder_id, new_name = None):
		try:
			file = service.files().get(fileId=file_id, fields='parents').execute()
			previous_parents = ",".join(file.get('parents'))
			if self.debug is False:
				file = self.service.files().update(fileId = file_id, addParents = folder_id, removeParents = previous_parents, fields = 'id, parents').execute()
				if new_name is not None:
					self.rename(file_id, new_name)
			return True
		except HttpError as error:
			if self.automatic is False:
				print(f"Script failed top move a file within google servers\n\nError: {error}")
			return False

	def rename(self, file_id, new_name):
		try:
			body = {'name': new_name}
			return service.files().update(fileId = file_id, body = body).execute()
		except HttpError as error:
			if self.automatic is False:
				print(f"Script failed to rename a document within google servers\n\nError: {error}")

class GmailConnection(System):
	def __init__(self, debug = False, automatic = False, log_file = None, working_directory = None):
		self.debug = debug
		self.automatic = automatic

		self.working_directory = working_directory

		if log_file != None:
			self.log_file = log_file

		with open('../credentials/gmail_creds.txt', 'r') as file:
			gmail_creds = file.read().split('\n')

		self.email = gmail_creds[0]
		self.password = gmail_creds[1]

	def send_email(self, to, subject, body, mime_type = 'plain', attachment = None):
		if self.debug == True:
			to = 'desc7849@colorado.edu'

		outlook = win32com.client.Dispatch('outlook.application')

		mail = outlook.CreateItem(0)
		mail.To = to
		mail.Subject = subject
		if mime_type == 'plain':
			mail.Body = body
		if mime_type == 'html':
			mail.HTMLBody = body

		if attachment:
			mail.Attachments.Add(Source=attachment) # Add the attachment pathway on computer too the mail

		if to in ['None', None, 'Yes']:
			return False
		try:
			mail.Send()
		except:
			print(f'Email failed to send to {to} - {subject}', file = self.log_file)


	def send_gmail(self, to, subject, body, mime_type = 'plain'):
		if self.debug == True:
			to = 'desc7849@colorado.edu'

		session = smtplib.SMTP_SSL('smtp-mail.outlook.com', 587)
		#session.starttls() # Enable extra security
		session.login(self.email, self.password)

		message = MIMEMultipart()
		message['From'] = self.email
		message['To'] = to
		message['Subject'] = subject
		message.attach(MIMEText(body, mime_type))
		text = message.as_string()

		session.sendmail(self.email, to, text)
		print(f"Email sent too {to}", file = self.log_file)

		session.quit()

class TwilioConnection(System):
	def __init__(self, debug = False, automatic = False, working_directory = None):
		self.debug = debug
		self.automatic = automatic

		self.working_directory = working_directory

		self.account_sid = os.environ.get('')
		self.auth_token = os.environ.get('')

		self.twilio_number = ''

		self.connect()
		atexit.register(self.disconnect)

	def __del__(self):
		self.disconnect()

	def __repr__(self):
		return f"Account SID: {self.account_sid}\nAuthorization Token: {self.auth_token}\nTwilio Number: {self.twilio_number}"

	def connect(self):
		try:
			self.client = Client(self.account_sid, self.auth_token)
			return True
		except:
			if self.automatic is False:
				print("Twilio connection failed")
			return False

	def disconnect(self):
		if self.client:
			del self.client

	def text(self, number, body):
		try:
			if debug is False:
				self.client.messages.create(from_ = os.environ.get(self.twilio_number), to = os.environ.get(number), body = body)
			elif self.automatic is False:
				print(f"Text '{body}' to be sent too {number}")
			return True
		except:
			if self.automatic is False:
				print(f"Text '{body}' to be sent too {number} failed too send")
			return False
