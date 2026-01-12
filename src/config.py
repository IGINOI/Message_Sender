import datetime

VERSION = "2.0.0" # Change this for every release
REPO = "IGINOI/Message_Sender"

# Set fake date for testing
# today = datetime.date.today() 
today = datetime.date(2025, 12, 29) 
# Range of days
range_of_days = 21


# global path to the appointements folder
appointments_file_path = f'C:/Users/david/Documents/CENTRO_FIT/APPUNTAMENTI'


# Perform backup of the appointments file
backup_enabled = False
# Send messages on WA
send_wamessage = True
# Shutdown the system after sending messages
shut_down = False
# Kill switch variable
continue_sending = True