import datetime
import sys
from pathlib import Path

VERSION = "2.3.1" # Change this for every release
REPO = "IGINOI/Message_Sender"

# Set fake date for testing
today = datetime.date.today() 
# today = datetime.date(2025, 12, 29) 
# Range of days
range_of_days = 21


# Detect if the application is running as a script or a frozen EXE
if getattr(sys, 'frozen', False):
    # Running as EXE: Use the folder containing the EXE
    base_dir = Path(sys.executable).parent
else:
    # Running as Script: Use the folder containing the .py file
    base_dir = Path(__file__).parent.parent

# Now, everything is relative to base_dir
# If the .ods file is in the SAME folder as the EXE:
appointments_file_path = base_dir


# Perform backup of the appointments file
backup_enabled = False
# Send messages on WA
send_wamessage = False
# Shutdown the system after sending messages
shut_down = False
# Kill switch variable
continue_sending = True