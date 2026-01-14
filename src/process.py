import pandas as pd
import datetime
import locale
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import os
import shutil

import config



def backup():
    source_path = f'{config.appointments_file_path}/Appuntamenti_{config.today.year}.xlsx'  # Replace with your file's source path
    destination_path = f'{config.appointments_file_path}/BACKUPS/Appuntamenti_{pd.Timestamp.now().strftime("%Y.%m.%d_%H.%M")}.xlsx'

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
    Appointments_file = pd.ExcelFile(f"{config.appointments_file_path}/Appuntamenti_{config.today.year}.xlsx")
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
            Processed_appointments[sheet_name] = pd.read_excel(Appointments_file, sheet_name=sheet_name, header=0)
        else:
            month_sheet = pd.read_excel(Appointments_file, sheet_name=sheet_name, header=0)
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

    for i in range(config.range_of_days):
        day = config.today + datetime.timedelta(days=i)
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
        if messagebox.askyesno("Conferma Spegnimento", "Vuoi spegnere il sistema dopo l'invio dei messaggi?", default=messagebox.NO):
            config.shut_down = True
        else:
            config.shut_down = False
        
        send_message()

    def send_message():
        # Save the selected indexes and close the GUI
        config.send_wamessage = True
        root.destroy()
    
    def print_messages():
        config.send_wamessage = False
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
    # print_button = ttk.Button(root, text="Stampa Messaggi", command=print_messages)
    # print_button.pack(pady=20)
    # send_button = ttk.Button(root, text="Invia Messaggi", command=request_system_shotdown)
    # send_button.pack(pady=20)

    # Put print_button and send_button side by side
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)
    print_button = ttk.Button(button_frame, text="Stampa Messaggi", command=print_messages)
    print_button.grid(row=0, column=0, padx=20)
    send_button = ttk.Button(button_frame, text="Invia Messaggi", command=request_system_shotdown)
    send_button.grid(row=0, column=1, padx=20)


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