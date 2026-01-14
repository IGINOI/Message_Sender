import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
from pyautogui import hotkey, press
from urllib.parse import quote
import pyautogui as pg
import time
import webbrowser as web

import config

WIDTH, HEIGHT = pg.size()

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
    
    for i, appointment in enumerate(appointments_with_contact):
        # CHECK IF STOP WAS CLICKED
        if not config.continue_sending:
            print("Invio interrotto dall'utente.")
            break

        # Logic for message content
        if appointment['Employee'] != 'Paola':
            wa_message = (f"Buongiorno {appointment['Customer'][0].split()[-1].capitalize()}, ricordiamo l'appuntamento di {appointment['Giorno']} {appointment['Mese']} alle {appointment['Ora']}.\nAttendiamo conferma, grazie.\nRicordiamo, per l'anno {config.today.year}, di portare l'impegnativa 'ciclo di massoterapia'.\nCentro Fit Roncegno Terme - Via Boschetti, 2.")
        else:
            wa_message = (f"Buongiorno {appointment['Customer'][0].split()[-1].capitalize()}, ricordiamo l'appuntamento di {appointment['Giorno']} {appointment['Mese']} alle {appointment['Ora']}.\nAttendiamo conferma, grazie.\nCentro Fit Roncegno Terme - Via Boschetti, 2.")

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

def print_message(appointment):
    with open("Verifica_Messaggi.txt", "a", encoding="utf-8") as f:
        f.write(f"{appointment['Customer'][0]}\n")
        f.write(f"{appointment['Giorno']} {appointment['Mese']} - {appointment['Ora']}\n")
        f.write(f"{appointment['Employee']}\n\n")

def send_whatsapp(phone_no: str, message: str, multiplier: int = 1) -> None:
    """Send WhatsApp Message"""

    # Open Web Page & wait for it to load
    web.open(f"https://web.whatsapp.com/send?phone={phone_no}&text={quote(message)}")
    time.sleep(25 * multiplier)
    
    # Select page & wait
    pg.click(WIDTH / 2 + 300, HEIGHT / 2)
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