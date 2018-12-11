import win32com.client
import os
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")	# Connect to Outlook using the current users account
folder = outlook.Folders['EMAIL_ADDRESS@DOMAIN.COM'].Folders['1ST LEVEL FOLDER NAME'] 				# Restricts the search to the specified folder within the main inbox. 
messages = folder.Items															# List of messages in the folder


def get_attachmets():
	"""This function will download all attachments from the specified folder in Outlook to the Current Working directory."""
	for i in range(messages.Count):
		print(messages[i].Subject)
		attachments = messages[i].Attachments
		subject = messages[i].Subject
		if attachments.Count == 0:
			pass
		else:
			for a in range(attachments.Count):
				attachment = attachments[a]
				attachment.SaveASFile(os.getcwd() + '\\' + str(attachment))
	
def delete_emails():
	"""This function will delete the emails after the attachments have been downloaded."""
	y = messages.Count
	for i in range(messages.Count):
		x = y - i - 1
		print(messages[x].Subject)
		messages[x].Delete()

get_attachmets()
delete_emails()
