import os
import config

from process import main, backup
from sender import show_progress_window
from check_updates import check_for_updates

if __name__ == "__main__":
    # Check for updates
    if check_for_updates():
        pass
    
    # Perform backup if needed
    if config.backup_enabled:
        backup()
    
    # Read the files and extract all the apoointments in the selected days
    appointments_with_contact = main()

    # Send the messages using pywhatkit
    show_progress_window(appointments_with_contact)
    
    # Shutdown the system if needed
    if config.shut_down:
        print("SUTTING DOWN SYSTEM")
        os.system('shutdown /s /t 5')