#! /usr/bin/python3
import sys
import smtplib
import SMTP_Credentials as creds
import logging
import datetime
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def start_logger():
	'''
	Starts the logging system.
	Returns:
		1) logger: Main Logger
		2) except_logger: Exceptions Logger
	'''
	# Create Logger
	logger = logging.getLogger('Main Logger')
	logger.setLevel(logging.INFO)

	# Create Exceptions Logger
	except_logger = logging.getLogger('Exceptions Logger')
	except_logger.setLevel(logging.ERROR)

	# create file handler and set level to info
	date = datetime.datetime.now()
	year = date.strftime('%Y')
	month = date.strftime('%B')
	try:
		file_handler = logging.FileHandler(os.path.join(os.getcwd(), 'Logs', year, f"{month}.log"))
	except FileNotFoundError:
		os.makedirs(f'Logs/{year}')
		file_handler = logging.FileHandler(os.path.join(os.getcwd(), 'Logs', year, f"{month}.log"))
	try:
		error_file = logging.FileHandler(os.path.join(os.getcwd(), 'Logs', year, 'Exceptions', f"{month}.log"))
	except FileNotFoundError:
		os.makedirs(f'Logs/{year}/Exceptions')
		error_file = logging.FileHandler(os.path.join(os.getcwd(), 'Logs', year, 'Exceptions', f"{month}.log"))

	# create formatter
	formatter = logging.Formatter("%(asctime)s [%(levelname)s]: send_email.py | %(message)s",
	                              "%Y-%m-%d %H:%M:%S")
	# add formatter to handlers
	file_handler.setFormatter(formatter)
	error_file.setFormatter(formatter)

	# add handlers to loggers
	logger.addHandler(file_handler)
	except_logger.addHandler(error_file)
	return logger, except_logger

def start_smtp(credentials):
	'''
	Starts the SMPT Server Connection.

	Requiers a credentials dictionary as:

	  credentials = {
	  	'name'     : SENDER NAME,
	  	'host'     : SMTP HOST DOMAIN,
	  	'port'     : SMTP PORT CONNECTION,
	  	'email'    : SENDER EMAIL,
	  	'password' : SENDER PASSWORD,
	  }
	'''
	#Start Connection
	print(f'Starting SMTP connection...\n{line}')
	smtp_connection = smtplib.SMTP(credentials['host'], credentials['port'])
	smtp_connection.ehlo()
	smtp_connection.starttls()
	smtp_connection.login(credentials['email'], credentials['password'])
	print(f'Successfully connected!\n{line}')
	return smtp_connection

def format_message(from_name, from_email, to_email, subject, text, html, reply_to = None):
	'''
	Formats the email content.

	Receives only string arguments.

	Requieres, in this exact order, the following arguments:
		from_name  = SENDER NAME 
		from_email = SENDER EMAIL
		to_email   = RECEPIENT EMAIL
		subject    = EMAIL SUBJECT
		text       = PLAIN TEXT OF THE EMAIL
		html       = HTML CODE OF THE EMAIL

	Optionally, an extra reply_to argument can be sent, which would be the Reply-To email of the message.
	'''
	print(f'formatting Message for {to_email}...\n{line}')
	# Create message container - the correct MIME type is multipart/alternative.
	msg = MIMEMultipart("alternative")
	msg["From"] = f'{from_name} <{from_email}>'
	msg["To"] = to_email
	msg["Subject"] = subject
	if reply_to:
		msg["Reply-to"] = reply_to

	# Record the MIME types of both parts - text/plain and text/html.
	part1 = MIMEText(text, "plain")
	part2 = MIMEText(html, "html")

	# Attach parts into message container.
	# According to RFC 2046, the last part of a multipart message, in this case
	# the HTML message, is best and preferred.
	msg.attach(part1)
	msg.attach(part2)
	return msg

def send_email(smtp_connection, from_name, from_email, user_input):
	'''
	Sends an email using a SMTP Connection.
	Requiers the following arguments:
	smtp_connectio = A SUCCESSFULLY SMTP CONNECTION (smtplib.SMTP object)
	from_name = SENDER NAME (String)
	from_email = SENDER EMAIL (String)
	user_input = USER INPUT (Tuple)
				 THIS TUPLE MUST CONTAIN THE FOLLOWING STRINGS:
					to_email   = RECEPIENT EMAIL
					subject    = EMAIL SUBJECT
					text       = PLAIN TEXT OF THE EMAIL
					html       = HTML CODE OF THE EMAIL
			        And, Optionally, an extra reply_to argument can be sent, which would be the Reply-To email of the message.
	'''
	message_details = (from_name, from_email) + user_input
	message = format_message(*message_details)
	print(f'Sending email...\n{line}')
	smtp_connection.sendmail(from_email, user_input[0], message.as_string())
	print(f'Message sent to: {user_input[0]}\n{line}')

def main(credentials, messages_list):
	#Start SMTP Connection
	smtp_connection = start_smtp(credentials)

	#Get Name and Email
	from_name  = credentials['name']
	from_email = credentials['email']

	#Send messages
	for user_input in messages_list:
		send_email(smtp_connection, from_name, from_email, user_input)

	#End SMTP Connection	
	smtp_connection.quit()

line = '-'*50
logger, except_logger = start_logger()

if __name__ == '__main__':
	error_message = """
To run this script, please send the following arguments:

	1) The SMTP Profile name (as stored in the SMTP Credentials File).
	2) A list (sent as a string) with the user's input.
		Ex: "[(to_email, subject, text, html, reply_to), ('ex@mple.com', 'test', 'Super cool message', 'Nice HTML Code')]

You can learn more using the help() command.
				"""	
	credentials_list = [profile for profile in dir(creds)[8:] if profile != 'test'] 
	profile = sys.argv[0].lower()

	if len(sys.argv) != 3:
		print(error_message)
	else:
		profile = sys.argv[1].lower()
		print(profile)
		if profile in credentials_list:
			exec(f'credentials = creds.{profile}')
		else:
			print(f'{profile} profile not found!\n{line}\nAborting...')
			print(error_message)
			exit()
		exec(f'messages_list = {sys.argv[2]}')
		main(credentials, messages_list)
		print("All emails Sent!")