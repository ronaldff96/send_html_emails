#! /usr/bin/python3
# =============================================================================
# Created  By  : Ronald Fonseca
# Creation Date: December 5, 2021
# =============================================================================
"""
This Script allows the user to send many HTML emails using SMTP Connection.
"""
# =============================================================================
# Imports
# =============================================================================
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
	try:
		print(f'Starting SMTP connection...\n{line}')
		smtp_connection = smtplib.SMTP(credentials['host'], credentials['port'])
		smtp_connection.ehlo()
		smtp_connection.starttls()
		smtp_connection.login(credentials['email'], credentials['password'])
	except Exception as error:
		logger.error(f"Error connecting to {credentials['email']}: {error}")
		except_logger.exception(f"Error connecting to {credentials['email']}: {error}")
		print(f"Error connecting to {credentials['email']}:")
		exit()
	else:
		print(f'Successfully connected!\n{line}')
		logger.info(f"Successfully connected to {credentials['email']}")
	return smtp_connection

def format_message(html_path, from_name, from_email, to_email, subject, text, html, reply_to = None):
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
	if html_path:
		with open(html, 'r', encoding='utf-8') as f:
			html_code = f.read()
		html = html_code
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

def send_email(smtp_connection, from_name, from_email, user_input, html_path):
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
					html       = HTML PATH / HTML CODE OF THE EMAIL
			        And, Optionally, an extra reply_to argument can be sent, which would be the Reply-To email of the message.
	html_path = BOOLEAN, TRUE IF HTML CONTENT IS SENT AS PATH OR FALSE IF IT IS SENT AS CODE.
	'''
	try:
		message_details = (from_name, from_email) + user_input
		message = format_message(html_path, *message_details)
		print(f'Sending email...\n{line}')
		smtp_connection.sendmail(from_email, user_input[0], message.as_string())
	except Exception as error:
		logger.error(f"Error sending message to {user_input[0]}: {error}")
		except_logger.exception(f"Error sending message to {user_input[0]}: {error}")
		print(f"Error sending message to {user_input[0]}:")
		exit()
	else:
		print(f'Message sent to {user_input[0]}\n{line}')
		logger.info(f'Message sent to {user_input[0]}')

def main(credentials, messages_list, html_path = True):
	#Start SMTP Connection
	smtp_connection = start_smtp(credentials)
	try:
		#Get Name and Email
		from_name  = credentials['name']
		from_email = credentials['email']

		#Send messages
		for user_input in messages_list:
			send_email(smtp_connection, from_name, from_email, user_input, html_path)

		#End SMTP Connection	
		smtp_connection.quit()
	except Exception as error:
		logger.error(f"Unexpected error: {error}")
		except_logger.exception(f"Unexpected error: {error}")
		print(f"Unexpected error ({error}), please check Logs for more info.")
		exit()

line = '-'*50
logger, except_logger = start_logger()

if __name__ == '__main__':
	error_message = """
To run this script, please send the following arguments:

	1) The SMTP Profile name (as stored in the SMTP Credentials File).
	2) A list (sent as a string) with the user's input.
		Ex: "[(to_email, subject, text, html, reply_to), ('ex@mple.com', 'test', 'Super cool message', 'HTML File Path')]
				"""	
	credentials_list = [profile for profile in dir(creds)[8:] if profile != 'test'] 
	profile = sys.argv[0].lower()

	if len(sys.argv) != 3:
		print(error_message)
	else:
		profile = sys.argv[1].lower()
		if profile in credentials_list:
			exec(f'credentials = creds.{profile}')
		else:
			print(f'{profile} profile not found!\n{line}\nAborting...')
			print(error_message)
			exit()
		exec(f'messages_list = {sys.argv[2]}')
		main(credentials, messages_list)
		print("All emails Sent!")