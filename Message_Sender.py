import pandas as pd
import datetime
import locale
import pywhatkit as kit
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import os
import shutil
import threading

# Set fake date for testing
#today = datetime.date.today() 
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


def backup():
    source_path = f'{appointments_file_path}/Appuntamenti_{today.year}.xlsx'  # Replace with your file's source path
    destination_path = f'{appointments_file_path}/BACKUPS/Appuntamenti_{pd.Timestamp.now().strftime("%Y.%m.%d_%H.%M")}.xlsx'

    # Ensure the source file exists
    if os.path.exists(source_path):
        try:
            # Copy the file to the new destination
            shutil.copy(source_path, destination_path)
        except Exception as e:
            print(f"An error occurred: {e}")
    else:
        print(f"Source file does not exist: {source_path}")


def main():
    # Read all sheets at once and preprocess the columns
    Appointments_file = pd.ExcelFile(f"{appointments_file_path}/Appuntamenti_{today.year}.xlsx")
    Appointments_file = preprocess_appointments(Appointments_file)

    # Extract the 7 days after today from the contacts sheet
    next_days_list = extract_days()
    next_days_list = start_gui(next_days_list)

    # Extract appointments for the selected days
    appointments_list, months_names = extract_appointments_for_next_days(Appointments_file, next_days_list)

    # Filter out empty appointments
    non_empty_appointments = filter_appointments(appointments_list, months_names)

    # Filter out appointments with a customer that does not have a phone numer in the contact list
    appointments_with_contact = filter_non_empty_appointments(Appointments_file, non_empty_appointments)

    return appointments_with_contact

def preprocess_appointments(Appointments_file):
    '''
    Preprocess the appointments file by duplicating days down the rows where the day is missing.
    '''
    sheets_name = Appointments_file.sheet_names

    Processed_appointments = {}

    for sheet_name in sheets_name:
        if sheet_name == "Contatti":
            Processed_appointments[sheet_name] = pd.read_excel(Appointments_file, sheet_name=sheet_name, header=1)
        else:
            month_sheet = pd.read_excel(Appointments_file, sheet_name=sheet_name, header=1)
            for i in range(len(month_sheet)-1):
                next_line = month_sheet.loc[i+1, "Giorno"]
                if pd.isna(next_line) or str(next_line).strip() == "":            
                    month_sheet.loc[(i+1), "Giorno"] = month_sheet.loc[i, "Giorno"]
            # Save back in Appointments_file the preprocessed month sheet
            Processed_appointments[sheet_name] = month_sheet  

    return Processed_appointments

def extract_days():
    locale.setlocale(locale.LC_TIME, "it-IT")
    days_list = []

    for i in range(range_of_days):
        day = today + datetime.timedelta(days=i)
        # %A -> weekday, %d -> day, %B -> month
        text = f"{day.strftime('%A')} {day.day} {day.strftime('%B')}" #day.strftime("%A %d %B")
        
        # Capitalize weekday AND month
        parts = text.split()
        parts[0] = parts[0].capitalize()
        parts[2] = parts[2].capitalize()
        
        days_list.append(" ".join(parts))
    return days_list

def start_gui(week_days_in_range):
    selected_indexes = []
    def on_toggle(index):
        # Toggle the index in the selected list
        adjusted_index = index + 1  # Adjust index to start from 1
        if adjusted_index in selected_indexes:
            selected_indexes.remove(adjusted_index)
            toggle_vars[index].set(0)  # Unselect the toggle button
        else:
            selected_indexes.append(adjusted_index)
            toggle_vars[index].set(1)  # Select the toggle button

    def request_system_shotdown():
        # ask whether at the end of the message sending the system has to be shutdown
        global shut_down
        if messagebox.askyesno("Conferma Spegnimento", "Vuoi spegnere il sistema dopo l'invio dei messaggi?", default=messagebox.NO):
            shut_down = True
        else:
            shut_down = False
        
        send_message()

    def send_message():
        # Save the selected indexes and close the GUI
        root.destroy()

    def on_close():
        # Ask the user for confirmation before closing
        if messagebox.askyesno("Conferma di uscita", "Sicuro di voler uscire?"):
            selected_indexes.clear()  # Clear the selected indexes
            root.destroy()  # Close the window

    root = tk.Tk()
    root.title("Selezione giorni della quale mandare avviso")
    root.protocol("WM_DELETE_WINDOW", on_close)

    # Create a style for the GUI
    style = ttk.Style()
    style.configure("TCheckbutton", font=("Helvetica", 12))
    style.configure("TButton", font=("Helvetica", 14))

    # Add a label at the top
    title_label = ttk.Label(root, text="Avvisa le persone con appuntamento nei seguenti giorni: ", font=("Helvetica", 14, "bold"))
    title_label.pack(pady=10)

    # Create a frame for the day selection grid
    grid_frame = ttk.Frame(root)
    grid_frame.pack(pady=10)

    # Create toggle buttons for each day in a grid layout
    toggle_vars = []
    for i, day in enumerate(week_days_in_range):
        var = tk.IntVar(value=0)  # 0 = off, 1 = on
        toggle_vars.append(var)
        toggle_button = ttk.Checkbutton(grid_frame, text=day, variable=var, command=lambda idx=i: on_toggle(idx))
        toggle_button.grid(row=i % 7, column=i // 7, padx=10, pady=5, sticky='w')  # Arrange in columns of 7

    # Add "Send Messages" button at the bottom
    send_button = ttk.Button(root, text="Invia Messaggi", command=request_system_shotdown)
    send_button.pack(pady=20)

    root.mainloop()

    weeks_days_selected = [week_days_in_range[i - 1] for i in selected_indexes]
    return weeks_days_selected

def extract_appointments_for_next_days(Appointments_file, next_days_list):
    filtered_appointments = []
    months_names = []
    for day in next_days_list:
        # Extract the name of the month
        month_name = day.split()[2]
        # Read the file of the selected month
        month_sheet = Appointments_file.get(month_name)
        # month_sheet.to_csv(f"Preprocessed_{month_name}_2_2026.csv", index=False)

        # Put the date in the wanted format
        day = str(day.split()[0:2]).replace("[", "").replace("]", "").replace("'", "").replace(",", "")
        months_names.append(month_name)
        filtered_appointments.append(month_sheet[month_sheet["Giorno"] == day])
    return filtered_appointments, months_names

def filter_appointments(appointments_list, months_names):
    employee_names = [("Marcus", "Ora"), ("Jacqueline", "Ora.1"), ("Eleonora", "Ora.2"), ("Paola", "Ora.3")]
    day_time_name_appointment = []

    for index, day in enumerate(appointments_list):
        for employee, hour in employee_names:
            for time in day[hour]:
                if time != "-" and not pd.isna(time):
                    # Extract the appointments matching timeslot and having as column name the selected employee
                    match = day[(day[hour] == time) & day[employee]]

                    # Format the time slots
                    time_str = str(time)
                    if isinstance(time, int):
                        formatted_time = time_str + ":00"
                    elif isinstance(time, float):
                        if len(time_str.split(".")[1]) == 2:
                            formatted_time = time_str.replace(".", ":")
                        elif len(time_str.split(".")[1]) == 1:
                            formatted_time = time_str.replace(".", ":") + "0"
                    
                    if not match.empty:
                        # Compose the appoitmnent entry: day, month, time, employee, customer
                        day_time_name_appointment.append({
                            "Giorno": day["Giorno"].values[0],
                            "Mese": months_names[index],
                            "Ora": formatted_time,
                            "Employee": employee,
                            "Customer": match[employee].values
                        })
    
    return day_time_name_appointment

def filter_non_empty_appointments(Appointments_file, appointments):
    # Read the contacts sheet (Processed_appointments is a dict of DataFrames)
    contacts_sheet = Appointments_file["Contatti"]

    # Build a lookup: "Nome Cognome" -> "Telefono"
    contacts_lookup = []
    for _, row in contacts_sheet.iterrows():
        nome = str(row.get("Nome")).strip()
        if nome == "nan":
            nome = ""

        cognome = str(row.get("Cognome")).strip()
        if cognome == "nan":
            cognome = ""
        
        full_name = (nome + " " + cognome).strip()
        telefono = str(row.get("Cellulare")).strip()
        contacts_lookup.append({
            "full_name": full_name,
            "Telephone": telefono })

    # Insert the telephone number in the appointments if present in contacts_lookup
    appointments_with_contact = []
    for appointment in appointments:
        customer_name = str(appointment["Customer"][0]).strip()
        # Search for the customer name in contacts_lookup
        contact_entry = next((item for item in contacts_lookup if item["full_name"] == customer_name), None)
        if contact_entry and contact_entry["Telephone"] and contact_entry["Telephone"] != "nan":
            # Add telephone to the appointment
            appointment["Telephone"] = contact_entry["Telephone"]
            appointments_with_contact.append(appointment)


    return appointments_with_contact


def show_progress_window(appointments_with_contact):
    global continue_sending
    continue_sending = True
    
    progress_root = tk.Tk()
    progress_root.title("Invio Messaggi")
    #progress_root.geometry("300x150")

    window_width = 300
    window_height = 150

    screen_width = progress_root.winfo_screenwidth()
    screen_height = progress_root.winfo_screenheight()
    
    # Calculate coordinates for bottom-right corner
    # (Subtracting a bit extra like 50px to account for the Taskbar)
    x = screen_width - window_width - 10
    y = screen_height - window_height - 60
    
    progress_root.geometry(f"{window_width}x{window_height}+{x}+0")
    progress_root.attributes("-topmost", True)

    label = ttk.Label(progress_root, text="Invio in corso...", font=("Helvetica", 10))
    label.pack(pady=20)

    def stop_process():
        global continue_sending
        if messagebox.askyesno("Stop", "Vuoi davvero interrompere l'invio?"):
            continue_sending = False
            label.config(text="Interruzione in corso...")

    stop_button = ttk.Button(progress_root, text="STOP INVIO", command=stop_process)
    stop_button.pack(pady=10)

    # Start the sending logic in a background thread so the GUI stays responsive
    threading.Thread(target=send_message_thread, args=(appointments_with_contact, progress_root, label), daemon=True).start()

    progress_root.mainloop()

def send_message_thread(appointments_with_contact, window, label_widget):
    global continue_sending
    
    for i, appointment in enumerate(appointments_with_contact):
        # CHECK IF STOP WAS CLICKED
        if not continue_sending:
            print("Invio interrotto dall'utente.")
            break

        # Logic for message content
        if appointment['Employee'] != 'Paola':
            wa_message = (f"\nBuongiorno {appointment['Customer'][0]}, ricordiamo l'appuntamento di {appointment['Giorno']} {appointment['Mese']} alle {appointment['Ora']}.\nAttendiamo conferma, grazie!\nRicordiamo, per chi non avesse ancora provveduto, di portare l'impegnativa 'ciclo di massoterapia' per l'anno {today.year}.\nCentro Fit Roncegno Terme - Via Boschetti, 2")
        else:
            wa_message = (f"\nBuongiorno {appointment['Customer'][0]}, ricordiamo l'appuntamento di {appointment['Giorno']} {appointment['Mese']} alle {appointment['Ora']}.\nAttendiamo conferma, grazie!\nCentro Fit Roncegno Terme - Via Boschetti, 2")

        # Update GUI Label
        label_widget.config(text=f"Invio {i+1} di {len(appointments_with_contact)}")
        
        if send_wamessage:
            kit.sendwhatmsg_instantly('+' + str(appointment["Telephone"]), wa_message, wait_time=20, tab_close=False)
        else:
            print(f"{wa_message}")
            import time
            time.sleep(0.5) # Simulating delay for testing

    window.destroy()



if __name__ == "__main__":
    # Perform backup if needed
    if backup_enabled:
        backup()
    
    # Read the files and extract all the apoointments in the selected days
    appointments_with_contact = main()

    # Send the messages using pywhatkit
    show_progress_window(appointments_with_contact)
    
    # Shutdown the system if needed
    if shut_down:
        print("SUTTING DOWN SYSTEM")
        os.system('shutdown /s /t 5')
