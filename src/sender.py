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
from config import today

WIDTH, HEIGHT = pg.size()

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

    stop_button = ttk.Button(progress_root, text="FERMA INVIO\n  MESSAGGI", command=stop_process)
    stop_button.pack(pady=10)
    # increase stop button size and font
    stop_button.config(width=15)

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
            wa_message = (f"\nBuongiorno {appointment['Customer'][0].split()[-1].capitalize()}, ricordiamo l'appuntamento di {appointment['Giorno']} {appointment['Mese']} alle {appointment['Ora']}.\nAttendiamo conferma, grazie!\n Ricordiamo, per l'anno {today.year}, di portare l'impegnativa 'ciclo di massoterapia'.\nCentro Fit Roncegno Terme - Via Boschetti, 2")
        else:
            wa_message = (f"\nBuongiorno {appointment['Customer'][0].split()[-1].capitalize()}, ricordiamo l'appuntamento di {appointment['Giorno']} {appointment['Mese']} alle {appointment['Ora']}.\nAttendiamo conferma, grazie!\nCentro Fit Roncegno Terme - Via Boschetti, 2")

        # Update GUI Label
        label_widget.config(text=f"Invio {i+1} di {len(appointments_with_contact)}")
        
        if config.send_wamessage:
            send_whatsapp('+' + str(appointment["Telephone"]), wa_message, wait_time=30, tab_close=True, close_time=10)
        else:
            print(f"{wa_message}")
            import time
            time.sleep(0.5) # Simulating delay for testing

    window.destroy()

 
def close_tab(wait_time: int = 2) -> None:
    """Closes the Currently Opened Browser Tab"""
    time.sleep(wait_time)
    hotkey("ctrl", "w")
    press("enter")

def check_number(number: str) -> bool:
    """Checks the Number to see if contains the Country Code"""

    return "+" in number or "_" in number

def send_whatsapp(phone_no: str, message: str, wait_time: int = 15, tab_close: bool = True, close_time: int = 3) -> None:
    """Send WhatsApp Message Instantly"""

    # If the number is not recognized, skip sending the message
    # if not check_number(number=phone_no):
    #     raise Exception("Country Code Missing in Phone Number!")

    # Open Web Page
    web.open(f"https://web.whatsapp.com/send?phone={phone_no}&text={quote(message)}")
    # Wait for the Page to load
    time.sleep(wait_time - 5)
    # Click on the screen to make it active
    pg.click(WIDTH / 2 + 300, HEIGHT / 2)
    # Wait to be sure the page is active
    time.sleep(5)
    # Press 'Enter' to send the message
    pg.press("enter")
    # Wait for the message to be sent
    time.sleep(5)
    # Log the message
    # log.log_message(_time=time.localtime(), receiver=phone_no, message=message)
    if tab_close:
        close_tab(wait_time = close_time)
    time.sleep(5)
    pg.press("enter")
    time.sleep(5)