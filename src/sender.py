import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
from urllib.parse import quote
import time
import webbrowser as web

import config

def show_progress_window(appointments_with_contact):
    
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
        if messagebox.askyesno("Stop", "Vuoi davvero interrompere l'invio?"):
            config.continue_sending = False
            label.config(text="Interruzione in corso...")

    stop_button = ttk.Button(progress_root, text="ANNULLA L'INVIO DEI\n MESSAGGI RESTANTI", command=stop_process)
    stop_button.pack(pady=10)
    # increase stop button size and font
    stop_button.config(width=22)

    # Start the sending logic in a background thread so the GUI stays responsive
    threading.Thread(target=send_message_thread, args=(appointments_with_contact, progress_root, label), daemon=True).start()

    progress_root.mainloop()

def send_message_thread(appointments_with_contact, window, label_widget):
    if config.shut_down:
        multiplier = 2
    else:
        multiplier = 1

    # Clear the verification file if it exists
    with open("Verifica_Messaggi.txt", "w", encoding="utf-8") as f:
        f.write(f"MESSAGGI DA MANDARE: {len(appointments_with_contact)}\n\n\n")
        f.write(f"ESEMPIO DI MESSAGGIO:\n\n")
        if config.custom_message:
            f.write(f"{config.custom_message}\n\n")
        else:
            f.write(f"{default_appointment_message(appointments_with_contact[0])}\n\n")
    
    for i, appointment in enumerate(appointments_with_contact):
        # CHECK IF STOP WAS CLICKED
        if not config.continue_sending:
            print("Invio interrotto dall'utente.")
            break

        # Use the GUI message when one was entered; otherwise keep the
        # appointment-specific default reminder.
        wa_message = config.custom_message or default_appointment_message(appointment)

        # Update GUI Label
        label_widget.config(text=f"Invio {i+1} di {len(appointments_with_contact)}")
        
        if config.send_wamessage:
            # Print the message in the file and send them via WhatsApp
            print_message(appointment)
            send_whatsapp('+' + str(appointment["Telephone"]), wa_message, multiplier=multiplier)
        else:
            # Print the messages in the file
            print_message(appointment)

    window.destroy()

def default_appointment_message(appointment):
    customer_last_name = appointment['Customer'][0].split()[-1].capitalize()
    appointment_details = (
        f"Buongiorno {customer_last_name}, ricordiamo l'appuntamento di "
        f"{appointment['Giorno']} {appointment['Mese']} alle {appointment['Ora']}.\n"
        "Attendiamo conferma, grazie.\n"
    )
    if appointment['Employee'] != 'Paola':
        appointment_details += (
            "Qualora non lo avesse già fatto, ricordiamo di portare l'impegnativa "
            "'ciclo di massoterapia' che dal 2026 ha una validità di 6 mesi.\n"
        )
    return appointment_details + "Centro Fit Roncegno Terme - Via Boschetti, 2."

def print_message(appointment):
    with open("Verifica_Messaggi.txt", "a", encoding="utf-8") as f:
        f.write(f"{appointment['Customer'][0]}\n")
        f.write(f"{appointment['Giorno']} {appointment['Mese']} - {appointment['Ora']}\n")
        f.write(f"{appointment['Telephone'][0:3]} {appointment['Telephone'][3:]}\n")
        f.write(f"{appointment['Employee']}\n\n")

def send_whatsapp(phone_no: str, message: str, multiplier: int = 1) -> None:
    """Send WhatsApp Message"""
    # pyautogui connects to the X display during import, so load it only
    # when a message is actually being sent.
    import pyautogui as pg
    from pyautogui import hotkey, press

    width, height = pg.size()

    # Open Web Page & wait for it to load
    web.open(f"https://web.whatsapp.com/send?phone={phone_no}&text={quote(message)}")
    time.sleep(25 * multiplier)
    
    # Select page & wait
    pg.click(width / 2 + 300, height / 2)
    time.sleep(10 * multiplier)
    
    # Press 'Enter' & wait for the message to be sent
    pg.press("enter")
    time.sleep(15 * multiplier)
    
    # Close tab and wait for it to close
    hotkey("ctrl", "w")
    press("enter")
    time.sleep(5 * multiplier)
    
    pg.press("enter")
    time.sleep(5 * multiplier)