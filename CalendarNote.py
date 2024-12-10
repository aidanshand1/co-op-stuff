import pickle
import os
from datetime import datetime, timedelta
import pytz
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from PIL import Image, ImageTk
import sys
import csv
from tkcalendar import Calendar

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

user_email = None
chosen_date_global = None

event_notes = {}    # event_id -> notes (string)
event_details = {}  # event_id -> {'summary': str, 'start_time': datetime, 'attendees': [str,...]}

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_token_path():
    home_dir = os.path.expanduser("~")
    token_dir = os.path.join(home_dir, ".calendar_app")
    os.makedirs(token_dir, exist_ok=True)
    return os.path.join(token_dir, "token.pickle")

def build_service():
    creds = None
    credentials_path = get_resource_path('credentials.json')
    token_path = get_token_path()

    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service

def parse_event_time(time_obj):
    if 'dateTime' in time_obj:
        return datetime.fromisoformat(time_obj['dateTime'])
    elif 'date' in time_obj:
        # All-day event
        return datetime.fromisoformat(time_obj['date'] + 'T00:00:00')
    return None

def get_events_for_date(calendar_id, chosen_date):
    service = build_service()
    atlantic = pytz.timezone('America/Halifax')
    chosen_date_local = atlantic.localize(chosen_date.replace(hour=0, minute=0, second=0, microsecond=0))

    start_of_day = chosen_date_local
    end_of_day = chosen_date_local + timedelta(days=1) - timedelta(microseconds=1)

    time_min_iso = start_of_day.isoformat()
    time_max_iso = end_of_day.isoformat()

    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=time_min_iso,
        timeMax=time_max_iso,
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    events = events_result.get('items', [])
    # Filter out "Home" or "Office"
    filtered_events = [e for e in events if e.get('summary', '') not in ['Home', 'Office']]
    return filtered_events

def open_notes_window(event_id):
    notes_window = tk.Toplevel(root)
    notes_window.title("Edit Notes")
    notes_window.geometry("600x400")  # width x height

    label = ttk.Label(notes_window, text="Enter your notes below:")
    label.pack(pady=5)

    text_box = tk.Text(notes_window, wrap=tk.WORD, width=80, height=20)
    text_box.pack(pady=5, fill=tk.BOTH, expand=True)

    existing_notes = event_notes.get(event_id, "")
    if existing_notes:
        text_box.insert(tk.END, existing_notes)

    def save_notes():
        new_notes = text_box.get("1.0", tk.END).strip()
        event_notes[event_id] = new_notes
        messagebox.showinfo("Notes Saved", "Your notes have been saved.")
        notes_window.destroy()

    save_button = ttk.Button(notes_window, text="Save Notes", command=save_notes)
    save_button.pack(pady=5)

def event_button_click(event_id):
    open_notes_window(event_id)

def show_events():
    if not user_email:
        messagebox.showwarning("Input Required", "Please enter your email address.")
        return
    if not chosen_date_global:
        messagebox.showwarning("Input Required", "Please select a date.")
        return

    try:
        events = get_events_for_date(user_email, chosen_date_global)

        for widget in events_frame.winfo_children():
            widget.destroy()

        if not events:
            no_event_label = ttk.Label(events_frame, text=f"No events found on {chosen_date_global.strftime('%A, %B %d, %Y')}.")
            no_event_label.pack(pady=10)
            return

        header_label = ttk.Label(events_frame, text=f"Events on {chosen_date_global.strftime('%A, %B %d, %Y')}:",
                                 font=('Arial', 14, 'bold'))
        header_label.pack(pady=(0, 10))

        event_details.clear()

        for event in events:
            summary = event.get('summary', 'No Title')
            start_time = parse_event_time(event.get('start', {}))
            end_time = parse_event_time(event.get('end', {}))
            attendees_list = event.get('attendees', [])
            attendees_emails = [a.get('email', '') for a in attendees_list]

            start_str = start_time.strftime('%I:%M %p') if start_time else 'N/A'
            end_str = end_time.strftime('%I:%M %p') if end_time else 'N/A'

            event_id = event.get('id', None)
            event_text = f"{summary} ({start_str} - {end_str})"
            event_button = ttk.Button(events_frame, text=event_text, command=lambda eid=event_id: event_button_click(eid))
            event_button.pack(pady=5, fill=tk.X)

            event_details[event_id] = {
                'summary': summary,
                'start_time': start_time,
                'attendees': attendees_emails
            }

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def copy_to_clipboard():
    if event_notes:
        notes_str = "Current Event Notes\n\n"
        for eid, notes in event_notes.items():
            details = event_details.get(eid, {})
            start_time = details.get('start_time', None)
            summary = details.get('summary', '')
            attendees = details.get('attendees', [])
            start_str = start_time.strftime('%Y-%m-%d %H:%M') if start_time else 'N/A'
            attendee_str = '; '.join(attendees)
            notes_str += (f"Event Start Time: {start_str}\n"
                          f"Event Name: {summary}\n"
                          f"Attendees: {attendee_str}\n"
                          f"Notes:\n{notes}\n\n"
                          "--------------------------------------------\n\n")
    else:
        notes_str = "No notes available."

    pyperclip.copy(notes_str)
    messagebox.showinfo("Copied", "Current notes copied to clipboard.")

def save_notes_to_csv():
    if not event_notes:
        messagebox.showinfo("No Notes", "There are no notes to save.")
        return

    if chosen_date_global:
        date_str = chosen_date_global.strftime('%Y_%m_%d')
    else:
        date_str = datetime.now().strftime('%Y_%m_%d')  # fallback if no chosen_date_global

    filename = f"event_notes_{date_str}.csv"
    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["Event Start Time", "Event Name", "Attendees", "Notes"])
        for eid, notes in event_notes.items():
            details = event_details.get(eid, {})
            start_time = details.get('start_time', None)
            summary = details.get('summary', '')
            attendees = details.get('attendees', [])
            start_str = start_time.strftime('%Y-%m-%d %H:%M') if start_time else 'N/A'
            attendee_str = '; '.join(attendees)
            writer.writerow([start_str, summary, attendee_str, notes])

    messagebox.showinfo("Saved", f"Notes have been saved to {filename}.")

def save_notes_to_txt():
    if not event_notes:
        messagebox.showinfo("No Notes", "There are no notes to save.")
        return

    if chosen_date_global:
        date_str = chosen_date_global.strftime('%Y_%m_%d')
    else:
        date_str = datetime.now().strftime('%Y_%m_%d')  # fallback if no chosen_date_global

    filename = f"event_notes_{date_str}.txt"
    with open(filename, 'w', encoding='utf-8') as f:
        f.write("Event Notes\n\n")
        for eid, notes in event_notes.items():
            details = event_details.get(eid, {})
            start_time = details.get('start_time', None)
            summary = details.get('summary', '')
            attendees = details.get('attendees', [])
            start_str = start_time.strftime('%Y-%m-%d %H:%M') if start_time else 'N/A'
            attendee_str = '; '.join(attendees)

            f.write(f"Event Start Time: {start_str}\n")
            f.write(f"Event Name: {summary}\n")
            f.write(f"Attendees: {attendee_str}\n")
            f.write("Notes:\n")
            f.write(f"{notes}\n\n")
            f.write("--------------------------------------------\n\n")

    messagebox.showinfo("Saved", f"Notes have been saved to {filename}.")

def pick_date():
    # Create a popup window to show the calendar
    date_window = tk.Toplevel(root)
    date_window.title("Select a Date")
    date_window.geometry("300x300")

    cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd')
    cal.pack(pady=20)

    def confirm_date():
        global chosen_date_global
        selected_date_str = cal.get_date()  # returns 'YYYY-MM-DD'
        chosen_date_global = datetime.strptime(selected_date_str, '%Y-%m-%d')
        date_window.destroy()
        # Update the date label on the main GUI
        if date_selected_label:
            date_selected_label.config(text=f"Selected Date: {chosen_date_global.strftime('%A, %B %d, %Y')}")

    confirm_button = ttk.Button(date_window, text="Confirm Date", command=confirm_date)
    confirm_button.pack(pady=10)

def initialize_app():
    global user_email
    user_email = email_entry.get().strip()
    if not user_email:
        messagebox.showwarning("Input Required", "Please enter your email address.")
        return
    email_entry_frame.destroy()
    display_main_gui()

def display_main_gui():
    global events_frame, date_selected_label

    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    header_label = ttk.Label(main_frame, text="Google Calendar Events Viewer", font=('Arial', 16, 'bold'))
    header_label.pack(pady=(0, 10))

    date_label = ttk.Label(main_frame, text="Select a Date:")
    date_label.pack(pady=5)

    pick_date_button = ttk.Button(main_frame, text="Open Calendar", command=pick_date)
    pick_date_button.pack(pady=5)

    # Label to show the chosen date
    date_selected_label = ttk.Label(main_frame, text="No date selected", font=('Arial', 12, 'italic'))
    date_selected_label.pack(pady=5)

    fetch_button = ttk.Button(main_frame, text="Show Events", command=show_events)
    fetch_button.pack(pady=10)

    copy_button = ttk.Button(main_frame, text="Copy Notes to Clipboard", command=copy_to_clipboard)
    copy_button.pack(pady=5)

    save_csv_button = ttk.Button(main_frame, text="Save All Notes to CSV", command=save_notes_to_csv)
    save_csv_button.pack(pady=5)

    save_txt_button = ttk.Button(main_frame, text="Save All Notes to TXT", command=save_notes_to_txt)
    save_txt_button.pack(pady=5)

    events_frame = ttk.Frame(main_frame)
    events_frame.pack(fill=tk.BOTH, expand=True, pady=10)

def create_gui():
    global root, email_entry_frame, email_entry
    root = tk.Tk()
    root.title("Google Calendar Events Viewer")

    window_width = 600
    window_height = 700
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

    style = ttk.Style()
    style.theme_use('clam')

    overlay_frame = tk.Frame(root, bg='white', bd=0, highlightthickness=0)
    overlay_frame.place(relx=0.5, rely=0.5, anchor='center')

    email_entry_frame = ttk.Frame(overlay_frame, padding="10")
    email_entry_frame.pack(pady=10)

    email_label = ttk.Label(email_entry_frame, text="Enter your email address:")
    email_label.pack(pady=5)

    email_entry = ttk.Entry(email_entry_frame, width=40)
    email_entry.pack(pady=5)

    submit_button = ttk.Button(email_entry_frame, text="Submit", command=initialize_app)
    submit_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
