import os

# Script to search for a certain file

def search(filename, directory = None):
	if directory == None:
		directory = os.getcwd()

	for content in os.listdir(directory):

		if os.path.isdir(content):
			results = search(filename, directory + '/' + content)
			if results == True:
				return results
		else:
			if content == filename:
				print(directory + '/' + filename)
				return True
	return False
