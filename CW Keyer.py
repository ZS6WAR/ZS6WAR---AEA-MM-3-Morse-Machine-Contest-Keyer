import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import serial.tools.list_ports
import serial
import win32com.client
import pythoncom
import time
import threading
from datetime import datetime, UTC
import queue
import json
import os
import webbrowser

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

# Global variables with defaults
SETTINGS_FILE = "settings.json"
ser = None
frequency_var = None
callsign_var = None
snt_var = None
rcv_var = None
exchange_var = None
speed_var = None
qso_list = []
serial_number = 1
log_tree = None
sent_text = None
frequency_queue = queue.Queue()
tune_state = None
macros = {
    "F1": "CQ CQ CQ DE {mycall} {mycall} K",
    "F2": "{callsign} {rst} {exchange}",
    "F3": "TU",
    "F4": "NR {mycall}",
    "F5": "{callsign} UR {rst} {exchange}",
    "F6": "NR?",
    "F7": "?",
    "F8": "",
    "F9": "",
    "F10": "",
    "F11": "",
    "F12": ""
}
button_labels = {
    "F1": "CQ",
    "F2": "Exchange",
    "F3": "TU",
    "F4": "My #",
    "F5": "His",
    "F6": "NR?",
    "F7": "?",
    "F8": "F8",
    "F9": "F9",
    "F10": "F10",
    "F11": "F11",
    "F12": "F12"
}
my_station = {
    "callsign": "ZS6WAR",
    "name": "",
    "address": "",
    "city": "",
    "country": "",
    "zipcode": "",
    "location": "",
    "cq_zone": "",
    "itu_zone": "",
    "rig": "",
    "antenna": "",
    "power": ""
}
contest_config = {
    "contest_name": "",
    "operator": "Single Op",
    "band": "All Bands",
    "power": "High",
    "transmitter": "One",
    "exchange": "",
    "use_serial_exchange": False
}

def load_settings():
    global qso_list, serial_number, macros, button_labels, my_station, contest_config
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                qso_list = settings.get("qso_list", [])
                serial_number = settings.get("serial_number", 1)
                macros.update(settings.get("macros", {}))
                button_labels.update(settings.get("button_labels", {}))
                my_station.update(settings.get("my_station", {}))
                contest_config.update(settings.get("contest_config", {}))
                return settings
        except Exception as e:
            pass  # Silently ignore errors
    return {}

def save_settings():
    settings = {
        "qso_list": qso_list,
        "serial_number": serial_number,
        "macros": macros,
        "button_labels": button_labels,
        "my_station": my_station,
        "contest_config": contest_config,
        "main_window_geometry": key_window.winfo_geometry() if 'key_window' in globals() else "900x550+0+0",
        "log_window_geometry": log_window.winfo_geometry() if 'log_window' in globals() else "800x400+0+0",
        "speed": speed_var.get() if speed_var else "25",
        "knob_mode": knob_mode.get() if 'knob_mode' in globals() else False,
        "sidetone_enabled": sidetone_enabled.get() if 'sidetone_enabled' in globals() else True,
        "repeat_enabled": repeat_enabled.get() if 'repeat_enabled' in globals() else False,
        "repeat_interval": repeat_interval.get() if 'repeat_interval' in globals() else "2.5",
        "use_5nn": use_5nn.get() if 'use_5nn' in globals() else False,
        "shorten_zeros": shorten_zeros.get() if 'shorten_zeros' in globals() else False
    }
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        pass  # Silently ignore errors

def show_port_selection():
    root = tk.Tk()
    root.title("ZS6WAR AEA MM-3 Morse Machine Contest Keyer - v.0.1 Setup")
    root.geometry("300x150")
    
    # Set window icon
    try:
        root.iconbitmap("morse_key.ico")
    except tk.TclError as e:
        pass  # Silently ignore icon errors

    def load_serial_ports():
        ports = [port.device for port in serial.tools.list_ports.comports()]
        if not ports:
            messagebox.showerror("Error", "No serial ports found!")
            port_var.set("No ports available")
            return ["No ports available"]
        port_var.set(ports[0])
        return ports

    def select_port():
        global ser
        selected_port = port_var.get()
        if selected_port == "No ports available":
            messagebox.showwarning("Warning", "Please connect a serial device!")
            return
        
        try:
            ser = serial.Serial(
                port=selected_port,
                baudrate=1200,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            messagebox.showinfo("Success", f"Connected to {selected_port} (8N1, 1200 baud)")
            root.destroy()
            show_function_key_window()
        except serial.SerialException as e:
            messagebox.showerror("Error", f"Failed to open port {selected_port}: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {e}")

    tk.Label(root, text="Select Serial Port:").pack(pady=10)
    port_var = tk.StringVar()
    port_dropdown = ttk.Combobox(root, textvariable=port_var, values=load_serial_ports(), state="readonly")
    port_dropdown.pack(pady=5)
    select_button = tk.Button(root, text="Select Port", command=select_port)
    select_button.pack(pady=10)

    root.mainloop()

def get_omnirig_data():
    while True:
        try:
            pythoncom.CoInitialize()
            omnirig = win32com.client.Dispatch("OmniRig.OmniRigX")
            time.sleep(2)
            
            rig = omnirig.Rig1
            status = rig.StatusStr if hasattr(rig, "StatusStr") else "Unknown"
            
            if status.lower() == "on-line":
                try:
                    rx_freq_hz = rig.GetRxFrequency()
                    if rx_freq_hz > 0:
                        freq_mhz = rx_freq_hz / 1_000_000
                        frequency_str = f"{freq_mhz:.6f} MHz"
                    else:
                        frequency_str = "N/A"
                except Exception:
                    frequency_str = "N/A"
            else:
                frequency_str = "N/A"
            
            frequency_queue.put(frequency_str)
            pythoncom.CoUninitialize()
        
        except Exception:
            frequency_queue.put("N/A")
            pythoncom.CoUninitialize()
        
        time.sleep(1)

def update_frequency():
    try:
        while not frequency_queue.empty():
            frequency_str = frequency_queue.get_nowait()
            frequency_var.set(frequency_str)
    except queue.Empty:
        pass
    key_window.after(100, update_frequency)

def calculate_cw_duration(message, wpm):
    dit_time = 1200 / wpm / 1000
    morse_units = {
        'A': 6, 'B': 10, 'C': 10, 'D': 8, 'E': 2, 'F': 8, 'G': 10, 'H': 6,
        'I': 4, 'J': 12, 'K': 10, 'L': 8, 'M': 10, 'N': 6, 'O': 12, 'P': 10,
        'Q': 12, 'R': 6, 'S': 4, 'T': 4, 'U': 6, 'V': 8, 'W': 8, 'X': 10,
        'Y': 12, 'Z': 10, '0': 14, '1': 12, '2': 10, '3': 8, '4': 6, '5': 4,
        '6': 6, '7': 8, '8': 10, '9': 12, ' ': 7
    }
    total_units = sum(morse_units.get(char.upper(), 3) for char in message)
    total_units += (len(message) - 1) * 1
    return total_units * dit_time + 2

def log_qso(event=None):
    global serial_number, log_tree
    callsign = callsign_var.get().strip()
    snt = snt_var.get().strip()
    rcv = rcv_var.get().strip()
    exchange_received = exchange_var.get().strip()
    frequency = frequency_var.get().strip()
    
    if not callsign or not exchange_received:
        messagebox.showwarning("Warning", "Callsign and Exchange are required to log a QSO!")
        return
    
    exchange_sent = str(serial_number) if contest_config["use_serial_exchange"] else contest_config["exchange"]
    
    qso = {
        "serial": serial_number,
        "datetime": datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S"),
        "callsign": callsign.upper(),
        "rst_sent": snt,
        "rst_received": rcv,
        "exchange_sent": exchange_sent,
        "exchange_received": exchange_received,
        "frequency": frequency.split()[0],
        "mode": "CW"
    }
    qso_list.append(qso)
    
    if log_tree:
        log_tree.insert("", "end", values=(len(qso_list), qso["datetime"], qso["callsign"], qso["rst_sent"], qso["rst_received"], qso["exchange_sent"], qso["exchange_received"], qso["frequency"], qso["mode"]))
    
    serial_number += 1
    callsign_var.set("")
    exchange_var.set("")
    callsign_entry.focus_set()

def show_qso_window():
    global log_tree
    qso_window = tk.Toplevel(key_window)
    qso_window.title("Logged QSOs")
    settings = load_settings()
    qso_window.geometry(settings.get("log_window_geometry", "800x400+0+0"))
    
    tree = ttk.Treeview(qso_window, columns=("Nr", "DateTime", "Callsign", "RST Sent", "RST Rcvd", "Exch Sent", "Exch Rcvd", "Freq", "Mode"), show="headings")
    tree.heading("Nr", text="Nr")
    tree.heading("DateTime", text="Date/Time (UTC)")
    tree.heading("Callsign", text="Callsign")
    tree.heading("RST Sent", text="RST Sent")
    tree.heading("RST Rcvd", text="RST Rcvd")
    tree.heading("Exch Sent", text="Exch Sent")
    tree.heading("Exch Rcvd", text="Exch Rcvd")
    tree.heading("Freq", text="Freq")
    tree.heading("Mode", text="Mode")
    
    for col in tree["columns"]:
        tree.column(col, width=90, anchor="center")
    
    for i, qso in enumerate(qso_list, 1):
        tree.insert("", "end", values=(i, qso["datetime"], qso["callsign"], qso["rst_sent"], qso["rst_received"], qso["exchange_sent"], qso["exchange_received"], qso["frequency"], qso["mode"]))
    
    context_menu = tk.Menu(tree, tearoff=0)
    context_menu.add_command(label="Delete", command=lambda: delete_qso(tree))
    context_menu.add_command(label="Edit", command=lambda: edit_qso(tree))

    def show_context_menu(event):
        if tree.identify_row(event.y):
            tree.selection_set(tree.identify_row(event.y))
            context_menu.post(event.x_root, event.y_root)

    tree.bind("<Button-3>", show_context_menu)

    tree.pack(fill="both", expand=True)
    log_tree = tree
    return qso_window

def delete_qso(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "No QSO selected!")
        return
    
    if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this QSO?"):
        serial = int(tree.item(selected_item, "values")[0])
        global qso_list
        qso_list.pop(serial - 1)
        tree.delete(*tree.get_children())
        for i, qso in enumerate(qso_list, 1):
            tree.insert("", "end", values=(i, qso["datetime"], qso["callsign"], qso["rst_sent"], qso["rst_received"], qso["exchange_sent"], qso["exchange_received"], qso["frequency"], qso["mode"]))

def edit_qso(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "No QSO selected!")
        return
    
    qso_values = tree.item(selected_item, "values")
    qso_number = int(qso_values[0])
    qso_index = qso_number - 1
    qso = qso_list[qso_index]

    dialog = tk.Toplevel(key_window)
    dialog.title(f"Edit QSO #{qso_number}")
    dialog.geometry("400x400")
    dialog.transient(key_window)
    dialog.grab_set()

    fields = [
        ("DateTime (UTC)", "datetime"),
        ("Callsign", "callsign"),
        ("RST Sent", "rst_sent"),
        ("RST Received", "rst_received"),
        ("Exchange Sent", "exchange_sent"),
        ("Exchange Received", "exchange_received"),
        ("Frequency", "frequency"),
        ("Mode", "mode")
    ]
    entries = {}

    for i, (label, key) in enumerate(fields):
        tk.Label(dialog, text=f"{label}:").grid(row=i, column=0, padx=5, pady=5, sticky="e")
        entry = tk.Entry(dialog, width=30)
        entry.insert(0, qso[key])
        entry.grid(row=i, column=1, padx=5, pady=5)
        entries[key] = entry

    def save_changes():
        for key, entry in entries.items():
            qso[key] = entry.get()
        qso_list[qso_index] = qso
        tree.delete(*tree.get_children())
        for i, qso in enumerate(qso_list, 1):
            tree.insert("", "end", values=(i, qso["datetime"], qso["callsign"], qso["rst_sent"], qso["rst_received"], qso["exchange_sent"], qso["exchange_received"], qso["frequency"], qso["mode"]))
        messagebox.showinfo("Updated", f"QSO #{qso_number} updated.")
        dialog.destroy()

    tk.Button(dialog, text="Save", command=save_changes).grid(row=len(fields), column=0, columnspan=2, pady=10)
    tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(fields)+1, column=0, columnspan=2, pady=5)

def export_to_adif():
    if not qso_list:
        messagebox.showinfo("Info", "No QSOs to export!")
        return
    
    filename = filedialog.asksaveasfilename(defaultextension=".adi", filetypes=[("ADIF Files", "*.adi"), ("All Files", "*.*")], title="Export to ADIF")
    if not filename:
        return
    
    with open(filename, "w") as f:
        f.write("Generated by ZS6WAR AEA MM-3 Morse Machine Contest Keyer - v.0.1\n<EOH>\n")
        for qso in qso_list:
            f.write(f"<QSO_DATE:8>{qso['datetime'][:10].replace('-', '')}\n")
            f.write(f"<TIME_ON:6>{qso['datetime'][11:].replace(':', '')}\n")
            f.write(f"<CALL:{len(qso['callsign'])}>{qso['callsign']}\n")
            f.write(f"<RST_SENT:{len(qso['rst_sent'])}>{qso['rst_sent']}\n")
            f.write(f"<RST_RCVD:{len(qso['rst_received'])}>{qso['rst_received']}\n")
            f.write(f"<STX:{len(qso['exchange_sent'])}>{qso['exchange_sent']}\n")
            f.write(f"<SRX:{len(qso['exchange_received'])}>{qso['exchange_received']}\n")
            f.write(f"<FREQ:{len(qso['frequency'])}>{qso['frequency']}\n")
            f.write("<MODE:2>CW\n")
            f.write("<EOR>\n")
    messagebox.showinfo("Success", f"Exported to {filename}")

def export_to_cabrillo():
    if not qso_list:
        messagebox.showinfo("Info", "No QSOs to export!")
        return
    
    filename = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Cabrillo Files", "*.log"), ("All Files", "*.*")], title="Export to Cabrillo")
    if not filename:
        return
    
    with open(filename, "w") as f:
        f.write("START-OF-LOG: 3.0\n")
        f.write(f"CALLSIGN: {my_station['callsign']}\n")
        f.write(f"CONTEST: {contest_config['contest_name']}\n")
        f.write(f"CATEGORY-OPERATOR: {contest_config['operator'].upper().replace(' ', '-')}\n")
        f.write(f"CATEGORY-BAND: {contest_config['band'].upper().replace(' ', '-')}\n")
        f.write(f"CATEGORY-POWER: {contest_config['power'].upper()}\n")
        f.write(f"CATEGORY-TRANSMITTER: {contest_config['transmitter'].upper()}\n")
        f.write("CREATED-BY: ZS6WAR AEA MM-3 Morse Machine Contest Keyer - v.0.1\n")
        for qso in qso_list:
            freq = int(float(qso["frequency"]) * 1000)
            date = qso["datetime"][:10].replace("-", "")
            time_str = qso["datetime"][11:].replace(":", "")[:4]
            f.write(f"QSO: {freq:5d} CW {date} {time_str} {my_station['callsign']:<13} {qso['rst_sent']} {qso['exchange_sent']}   {qso['callsign']:<13} {qso['rst_received']} {qso['exchange_received']}\n")
        f.write("END-OF-LOG:\n")
    messagebox.showinfo("Success", f"Exported to {filename}")

def read_serial():
    global ser, sent_text, last_sent_complete
    while True:
        try:
            if ser and ser.is_open:
                line = ser.readline().decode('ascii', errors='ignore').strip()
                if line:
                    sent_text.config(state="normal")
                    sent_text.delete(1.0, tk.END)
                    sent_text.insert(tk.END, line)
                    sent_text.config(state="disabled")
                    sent_text.see(tk.END)
                    last_sent_complete.set()
        except serial.SerialException:
            break
        except Exception:
            pass
        time.sleep(0.1)

def show_contest_setup():
    global tune_state
    if not tune_state.get():
        dialog = tk.Toplevel(key_window)
        dialog.title("Contest Setup")
        dialog.geometry("300x450")
        dialog.transient(key_window)
        dialog.grab_set()

        tk.Label(dialog, text="Contest Name:").pack(pady=5)
        contest_name_entry = tk.Entry(dialog, width=30)
        contest_name_entry.insert(0, contest_config["contest_name"])
        contest_name_entry.pack(pady=5)

        tk.Label(dialog, text="Operator:").pack(pady=5)
        operator_var = tk.StringVar(value=contest_config["operator"])
        operator_menu = ttk.Combobox(dialog, textvariable=operator_var, values=["Single Op", "Multi Op"], state="readonly")
        operator_menu.pack(pady=5)

        tk.Label(dialog, text="Band:").pack(pady=5)
        band_var = tk.StringVar(value=contest_config["band"])
        band_menu = ttk.Combobox(dialog, textvariable=band_var, values=["160M", "80M", "40M", "20M", "15M", "10M", "All Bands"], state="readonly")
        band_menu.pack(pady=5)

        tk.Label(dialog, text="Power:").pack(pady=5)
        power_var = tk.StringVar(value=contest_config["power"])
        power_menu = ttk.Combobox(dialog, textvariable=power_var, values=["High", "Low", "QRP"], state="readonly")
        power_menu.pack(pady=5)

        tk.Label(dialog, text="Transmitter:").pack(pady=5)
        transmitter_var = tk.StringVar(value=contest_config["transmitter"])
        transmitter_menu = ttk.Combobox(dialog, textvariable=transmitter_var, values=["One", "Two"], state="readonly")
        transmitter_menu.pack(pady=5)

        tk.Label(dialog, text="Exchange:").pack(pady=5)
        
        use_serial_var = tk.BooleanVar(value=contest_config["use_serial_exchange"])
        serial_check = tk.Checkbutton(dialog, text="Use Serial Number", variable=use_serial_var)
        serial_check.pack(pady=5)

        exchange_entry = tk.Entry(dialog, width=30)
        exchange_entry.insert(0, contest_config["exchange"])
        exchange_entry.pack(pady=5)

        def toggle_exchange_entry():
            if use_serial_var.get():
                exchange_entry.config(state="disabled")
            else:
                exchange_entry.config(state="normal")

        toggle_exchange_entry()
        serial_check.config(command=toggle_exchange_entry)

        def save_contest_config():
            contest_config["contest_name"] = contest_name_entry.get()
            contest_config["operator"] = operator_var.get()
            contest_config["band"] = band_var.get()
            contest_config["power"] = power_var.get()
            contest_config["transmitter"] = transmitter_var.get()
            contest_config["exchange"] = exchange_entry.get()
            contest_config["use_serial_exchange"] = use_serial_var.get()
            messagebox.showinfo("Updated", "Contest configuration saved.")
            dialog.destroy()

        tk.Button(dialog, text="Save", command=save_contest_config).pack(pady=10)
        tk.Button(dialog, text="Cancel", command=dialog.destroy).pack(pady=5)

def show_function_key_window():
    global key_window, freq_label, frequency_var, callsign_var, snt_var, rcv_var, exchange_var, speed_var, log_tree, sent_text, tune_state
    global knob_mode, sidetone_enabled, repeat_enabled, repeat_interval, use_5nn, shorten_zeros, log_window, callsign_entry, keyboard_keyer
    
    settings = load_settings()
    
    key_window = tk.Tk()
    key_window.title("ZS6WAR AEA MM-3 Morse Machine Contest Keyer - v.0.1")
    key_window.geometry(settings.get("main_window_geometry", "900x550+0+0"))

    try:
        key_window.iconbitmap("morse_key.ico")
    except tk.TclError:
        pass  # Silently ignore icon errors

    frequency_var = tk.StringVar(value="N/A")
    callsign_var = tk.StringVar()
    snt_var = tk.StringVar(value="599")
    rcv_var = tk.StringVar(value="599")
    exchange_var = tk.StringVar(value="")
    tune_state = tk.BooleanVar(value=False)
    speed_var = tk.StringVar(value=settings.get("speed", "25"))
    use_5nn = tk.BooleanVar(value=settings.get("use_5nn", False))
    shorten_zeros = tk.BooleanVar(value=settings.get("shorten_zeros", False))
    knob_mode = tk.BooleanVar(value=settings.get("knob_mode", False))
    sidetone_enabled = tk.BooleanVar(value=settings.get("sidetone_enabled", True))
    repeat_enabled = tk.BooleanVar(value=settings.get("repeat_enabled", False))
    repeat_interval = tk.StringVar(value=settings.get("repeat_interval", "2.5"))

    ui_elements = []

    log_window = show_qso_window()
    log_window.protocol("WM_DELETE_WINDOW", lambda: None)
    
    def format_output(callsign, rst, received_exchange, mycall):
        formatted_rst = "5NN" if use_5nn.get() and rst == "599" else rst
        formatted_received_exchange = received_exchange
        if shorten_zeros.get() and received_exchange.isdigit():
            formatted_received_exchange = received_exchange.replace("0", "T")
        formatted_exchange = str(serial_number) if contest_config["use_serial_exchange"] else contest_config["exchange"]
        formatted_serial = str(serial_number)
        return callsign, formatted_rst, formatted_received_exchange, mycall, formatted_exchange, formatted_serial

    def send_to_serial(message):
        global ser, serial_number, speed_var
        if ser is None or not ser.is_open:
            messagebox.showerror("Error", "Serial port not connected. Please restart and select a port.")
            key_window.destroy()
            show_port_selection()
            return 0
        
        callsign, rst, received_exchange, mycall, exchange, serial = format_output(
            callsign_var.get().strip(),
            snt_var.get().strip(),
            exchange_var.get().strip(),
            my_station["callsign"]
        )
        formatted_message = message.format(
            callsign=callsign,
            rst=rst,
            exchange=exchange,
            mycall=mycall,
            serial=serial
        )
        try:
            ser.write(formatted_message.encode('ascii') + b'\n')
            time.sleep(0.1)
            ser.reset_input_buffer()
            wpm = int(speed_var.get())
            duration = calculate_cw_duration(formatted_message, wpm)
            return duration
        except serial.SerialException as e:
            messagebox.showerror("Error", f"Failed to send: {e}")
            ser.close()
            ser = None
            key_window.destroy()
            show_port_selection()
            return 0
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid WPM value: {speed_var.get()}")
            return 0
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {e}")
            return 0

    def send_keyer_text(event):
        char = event.char
        if char and ser and ser.is_open and not tune_state.get():
            try:
                ser.write(char.encode('ascii'))
            except Exception:
                pass

    def send_command(command, include_terminator=True, tune_command=False):
        global ser
        try:
            ser.write("\x03".encode('ascii') + b'\n')
            ser.write(command.encode('ascii') + b'\n')
            if include_terminator:
                terminator = "*9" if tune_command else "*C709"
                ser.write(terminator.encode('ascii') + b'\n')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send command: {e}")

    def set_knob_mode():
        if not tune_state.get():
            command = "*B6" if knob_mode.get() else "*A6"
            send_command(command)
            if not knob_mode.get():
                set_speed()

    def set_speed():
        if not tune_state.get() and not knob_mode.get():
            try:
                speed = int(speed_var.get())
                if 1 <= speed <= 99:
                    command = f"*6{speed:02d}"
                    send_command(command)
                else:
                    speed_var.set("25")
            except ValueError:
                speed_var.set("25")

    def set_sidetone():
        if not tune_state.get():
            command = "*A1" if sidetone_enabled.get() else "*B1"
            send_command(command)

    def toggle_tune():
        if tune_state.get():
            send_command("", include_terminator=True, tune_command=True)
            for element in ui_elements:
                if element != speed_entry:
                    element.config(state="normal")
                else:
                    element.config(state="readonly")
            tune_state.set(False)
            knob_mode.set(True)
            set_knob_mode()
            time.sleep(0.1)
            knob_mode.set(False)
            set_knob_mode()
        else:
            send_command("*3**7", include_terminator=False, tune_command=True)
            for element in ui_elements:
                element.config(state="disabled")
            tune_button.config(state="normal")
            tune_state.set(True)

    def increase_speed(event=None):
        if not tune_state.get() and not knob_mode.get():
            try:
                speed = int(speed_var.get())
                if speed < 99:
                    speed_var.set(str(speed + 1))
                    set_speed()
            except ValueError:
                speed_var.set("25")
                set_speed()

    def decrease_speed(event=None):
        if not tune_state.get() and not knob_mode.get():
            try:
                speed = int(speed_var.get())
                if speed > 1:
                    speed_var.set(str(speed - 1))
                    set_speed()
            except ValueError:
                speed_var.set("25")
                set_speed()

    repeating = False
    last_sent_complete = threading.Event()

    def repeat_cq():
        nonlocal repeating
        while repeating and not tune_state.get() and not knob_mode.get():
            last_sent_complete.clear()
            duration = send_to_serial(macros["F1"])
            time.sleep(duration)
            last_sent_complete.set()
            try:
                interval = float(repeat_interval.get())
                if interval <= 0:
                    interval = 2.5
            except ValueError:
                interval = 2.5
            time.sleep(interval)
        repeating = False

    def start_repeat():
        nonlocal repeating
        if not tune_state.get() and not knob_mode.get() and not repeating:
            repeating = True
            threading.Thread(target=repeat_cq, daemon=True).start()
        elif knob_mode.get():
            messagebox.showwarning("Warning", "Repeat function is disabled when Knob mode is selected.")
            repeating = False

    def stop_repeat():
        nonlocal repeating
        repeating = False

    def stop_repeat_and_untick(event=None):
        nonlocal repeating
        repeating = False
        repeat_enabled.set(False)

    def create_action(f_key):
        def action():
            if not tune_state.get():
                message = macros[f_key]
                if message:
                    if f_key == "F1" and repeat_enabled.get():
                        start_repeat()
                        return
                    if f_key == "F2":
                        if not snt_var.get().strip():
                            snt_var.set("599")
                        if not rcv_var.get().strip():
                            rcv_var.set("599")
                        exchange_entry.focus_set()
                    send_to_serial(message)
        return action

    def lookup_qrz():
        callsign = callsign_var.get().strip()
        if not callsign:
            messagebox.showwarning("Warning", "Please enter a callsign to look up!")
            return
        url = f"https://www.qrz.com/db/{callsign.upper()}"
        webbrowser.open(url)

    def edit_macro_and_label(f_key):
        if not tune_state.get():
            dialog = tk.Toplevel(key_window)
            dialog.title(f"Edit {f_key}")
            dialog.geometry("400x300")
            dialog.transient(key_window)
            dialog.grab_set()

            tk.Label(dialog, text="Label (button text below F-key):").pack(pady=5)
            label_entry = tk.Entry(dialog, width=50)
            label_entry.insert(0, button_labels[f_key])
            label_entry.pack(pady=5)

            tk.Label(dialog, text="Macro (message to send):").pack(pady=5)
            macro_entry = tk.Entry(dialog, width=50)
            macro_entry.insert(0, macros[f_key])
            macro_entry.pack(pady=5)

            tk.Label(dialog, text="Insert Placeholder:").pack(pady=5)
            placeholder_var = tk.StringVar()
            placeholders = {
                "hiscall": "{callsign}",
                "mycall": "{mycall}",
                "hisrst": "{rst}",
                "myrst": "599",
                "exchange": "{exchange}",
                "serial": "{serial}"
            }
            placeholder_menu = ttk.Combobox(dialog, textvariable=placeholder_var, values=list(placeholders.keys()), state="readonly")
            placeholder_menu.pack(pady=5)

            def insert_placeholder():
                selected = placeholder_var.get()
                if selected:
                    insert_text = placeholders[selected]
                    cursor_pos = macro_entry.index(tk.INSERT)
                    macro_entry.insert(cursor_pos, insert_text)
                    macro_entry.focus_set()

            tk.Button(dialog, text="Insert", command=insert_placeholder).pack(pady=5)

            def save_changes():
                new_macro = macro_entry.get()
                new_label = label_entry.get()
                if new_macro is not None and new_label is not None:
                    macros[f_key] = new_macro
                    button_labels[f_key] = new_label
                    buttons[f_key].config(text=f"{f_key}\n{new_label}")
                    messagebox.showinfo("Updated", f"{f_key} updated:\nLabel: '{new_label}'\nMacro: '{new_macro}'")
                dialog.destroy()

            tk.Button(dialog, text="Save", command=save_changes).pack(pady=10)
            tk.Button(dialog, text="Cancel", command=dialog.destroy).pack(pady=5)

    def show_my_station():
        if not tune_state.get():
            dialog = tk.Toplevel(key_window)
            dialog.title("My Station")
            dialog.geometry("400x500")
            dialog.transient(key_window)
            dialog.grab_set()

            entries = {}
            station_items = [
                ("callsign", "My Callsign"),
                ("name", "My Name"),
                ("address", "My Address"),
                ("city", "My City"),
                ("country", "My Country"),
                ("zipcode", "My Zipcode"),
                ("location", "My Location (Gridsquare)"),
                ("cq_zone", "My CQ Zone"),
                ("itu_zone", "My ITU Zone"),
                ("rig", "My Rig"),
                ("antenna", "My Antenna"),
                ("power", "My Power Setting")
            ]
            
            for i, (key, name) in enumerate(station_items):
                tk.Label(dialog, text=f"{name}:").grid(row=i, column=0, padx=5, pady=5, sticky="e")
                entry = tk.Entry(dialog, width=40)
                entry.insert(0, my_station[key])
                entry.grid(row=i, column=1, padx=5, pady=5)
                entries[key] = entry

            def save_changes():
                for key, entry in entries.items():
                    my_station[key] = entry.get()
                messagebox.showinfo("Updated", "My Station details saved.")
                dialog.destroy()

            tk.Button(dialog, text="Save", command=save_changes).grid(row=len(station_items), column=0, columnspan=2, pady=10)
            tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(station_items)+1, column=0, columnspan=2, pady=5)

    def show_shorten_chars():
        if not tune_state.get():
            settings_window = tk.Toplevel(key_window)
            settings_window.title("Shorten Characters")
            settings_window.geometry("300x150")
            tk.Label(settings_window, text="RST Format:").pack(pady=5)
            tk.Radiobutton(settings_window, text="599", variable=use_5nn, value=False).pack()
            tk.Radiobutton(settings_window, text="5NN", variable=use_5nn, value=True).pack()
            tk.Label(settings_window, text="Serial Numbers:").pack(pady=5)
            tk.Checkbutton(settings_window, text="Shorten 0 to T", variable=shorten_zeros).pack()
            tk.Button(settings_window, text="Close", command=settings_window.destroy).pack(pady=10)

    def start_new_contest():
        if messagebox.askyesno("New Contest", "Are you sure you want to start a new contest?\nThis will delete the current log."):
            global qso_list, serial_number, log_tree
            qso_list = []
            serial_number = 1
            if log_tree:
                for item in log_tree.get_children():
                    log_tree.delete(item)
            messagebox.showinfo("New Contest", "New contest started. Log has been cleared.")
            save_settings()

    def on_closing():
        global ser
        nonlocal repeating
        repeating = False
        save_settings()
        if ser is not None and ser.is_open:
            ser.close()
        key_window.destroy()

    menubar = tk.Menu(key_window)
    key_window.config(menu=menubar)
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="New Contest", command=start_new_contest)
    file_menu.add_command(label="My Station", command=show_my_station)
    file_menu.add_command(label="Show QSOs", command=show_qso_window)
    file_menu.add_command(label="Export to ADIF", command=export_to_adif)
    file_menu.add_command(label="Export to Cabrillo", command=export_to_cabrillo)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=on_closing)

    contesting_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Contesting", menu=contesting_menu)
    contesting_menu.add_command(label="Contest Setup", command=show_contest_setup)
    macros_menu = tk.Menu(contesting_menu, tearoff=0)
    contesting_menu.add_cascade(label="Macros", menu=macros_menu)
    for f_key in macros.keys():
        macros_menu.add_command(label=f"Edit {f_key}", command=lambda k=f_key: edit_macro_and_label(k))
    contesting_menu.add_command(label="Shorten Chars", command=show_shorten_chars)

    freq_frame = tk.Frame(key_window)
    freq_frame.grid(row=0, column=0, columnspan=4, pady=5, sticky="w")
    freq_label = tk.Label(freq_frame, text="Freq: ", textvariable=frequency_var, font=("Arial", 18))
    freq_label.pack(side="left", padx=5)
    ui_elements.append(freq_label)

    time_frame = tk.Frame(key_window)
    time_frame.grid(row=0, column=4, columnspan=4, pady=5, sticky="e")
    time_label = tk.Label(time_frame, text="UTC: --:--:--", font=("Arial", 18))
    time_label.pack(side="right", padx=5)
    ui_elements.append(time_label)

    def update_time():
        utc_time = datetime.now(UTC).strftime("%H:%M:%S")
        time_label.config(text=f"UTC: {utc_time}")
        key_window.after(1000, update_time)

    callsign_label = tk.Label(key_window, text="Callsign:")
    callsign_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
    callsign_entry = tk.Entry(key_window, textvariable=callsign_var, width=15)
    callsign_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Label(key_window, text="SNT:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
    snt_entry = tk.Entry(key_window, textvariable=snt_var, width=5)
    snt_entry.grid(row=1, column=3, padx=5, pady=5)
    tk.Label(key_window, text="RCV:").grid(row=1, column=4, padx=5, pady=5, sticky="e")
    rcv_entry = tk.Entry(key_window, textvariable=rcv_var, width=5)
    rcv_entry.grid(row=1, column=5, padx=5, pady=5)
    tk.Label(key_window, text="Exchange:").grid(row=1, column=6, padx=5, pady=5, sticky="e")
    exchange_entry = tk.Entry(key_window, textvariable=exchange_var, width=15)
    exchange_entry.grid(row=1, column=7, padx=5, pady=5)
    ui_elements.append(callsign_label)

    callsign_entry.bind('<KeyPress>', lambda event: stop_repeat())

    buttons = {}
    for i in range(1, 13):
        f_key = f"F{i}"
        row = 2 if i <= 6 else 3
        col = (i-1) % 6
        btn = tk.Button(key_window, text=f"{f_key}\n{button_labels[f_key]}", command=create_action(f_key), width=8, height=2)
        btn.grid(row=row, column=col, padx=10, pady=10)
        buttons[f_key] = btn
        ui_elements.append(btn)

    qrz_button = tk.Button(key_window, text="QRZ", command=lookup_qrz, width=8, height=2)
    qrz_button.grid(row=2, column=6, padx=10, pady=10, rowspan=2)
    ui_elements.append(qrz_button)

    control_panel = tk.Frame(key_window)
    control_panel.grid(row=4, column=0, columnspan=8, pady=5)

    speed_frame = tk.Frame(control_panel)
    speed_frame.pack(side="left", padx=5, pady=5)
    tk.Label(speed_frame, text="Speed:").pack(side="top", pady=2)
    knob_check = tk.Checkbutton(speed_frame, text="Knob", variable=knob_mode, command=set_knob_mode)
    knob_check.pack(side="left", padx=2, pady=5)
    ui_elements.append(knob_check)
    speed_entry = tk.Entry(speed_frame, textvariable=speed_var, width=5, state="readonly")
    speed_entry.pack(side="left", padx=2, pady=5)
    ui_elements.append(speed_entry)
    tk.Label(speed_frame, text="WPM").pack(side="left", padx=2, pady=5)

    control_frame = tk.Frame(control_panel)
    control_frame.pack(side="left", padx=5, pady=5)
    tk.Label(control_frame, text="Control:").pack(side="top", pady=2)
    sidetone_check = tk.Checkbutton(control_frame, text="Sidetone", variable=sidetone_enabled, command=set_sidetone)
    sidetone_check.pack(side="left", padx=2, pady=5)
    ui_elements.append(sidetone_check)
    tune_button = tk.Button(control_frame, text="Tune", command=toggle_tune)
    tune_button.pack(side="left", padx=2, pady=5)
    ui_elements.append(tune_button)

    repeat_frame = tk.Frame(control_panel)
    repeat_frame.pack(side="left", padx=5, pady=5)
    tk.Label(repeat_frame, text="Repeat:").pack(side="top", pady=2)
    repeat_check = tk.Checkbutton(repeat_frame, variable=repeat_enabled)
    repeat_check.pack(side="left", padx=1, pady=5)
    ui_elements.append(repeat_check)
    repeat_entry = tk.Entry(repeat_frame, textvariable=repeat_interval, width=5)
    repeat_entry.pack(side="left", padx=1, pady=5)
    ui_elements.append(repeat_entry)
    tk.Label(repeat_frame, text="sec").pack(side="left", padx=2, pady=5)

    sent_frame = tk.Frame(key_window)
    sent_frame.grid(row=5, column=0, columnspan=8, pady=5, sticky="ew")
    tk.Label(sent_frame, text="Sent:").pack(side="left", padx=5)
    sent_text = tk.Text(sent_frame, height=1, width=50, state="disabled")
    sent_text.pack(side="left", fill="x", expand=True, padx=5)
    ui_elements.append(sent_text)

    keyer_frame = tk.Frame(key_window)
    keyer_frame.grid(row=6, column=0, columnspan=8, pady=5, sticky="ew")
    tk.Label(keyer_frame, text="Keyboard Keyer:").pack(side="left", padx=5)
    keyboard_keyer = tk.Text(keyer_frame, height=1, width=50)
    keyboard_keyer.pack(side="left", fill="x", expand=True, padx=5)
    keyboard_keyer.bind("<KeyPress>", send_keyer_text)
    ui_elements.append(keyboard_keyer)

    key_window.bind('<Prior>', increase_speed)
    key_window.bind('<Next>', decrease_speed)
    key_window.bind('<Return>', log_qso)
    key_window.bind('<Escape>', stop_repeat_and_untick)

    for col in range(8):
        key_window.grid_columnconfigure(col, weight=1, uniform="btn_group")

    for i in range(1, 13):
        f_key = f"F{i}"
        key_window.bind(f'<F{i}>', lambda e, k=f_key: create_action(k)())

    threading.Thread(target=get_omnirig_data, daemon=True).start()
    threading.Thread(target=read_serial, daemon=True).start()

    set_knob_mode()
    set_speed()
    set_sidetone()
    update_frequency()
    update_time()

    key_window.protocol("WM_DELETE_WINDOW", on_closing)
    key_window.mainloop()

if __name__ == "__main__":
    load_settings()
    show_port_selection()
