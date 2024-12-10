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
import random

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

TIMEZONES = {
    "Atlantic Standard Time": 'America/Halifax',
    "Eastern Standard Time": 'America/New_York',
    "Central Standard Time": 'America/Chicago',
    "Mountain Standard Time": 'America/Denver',
    "Pacific Standard Time": 'America/Los_Angeles',
    "UTC": 'UTC',
    # Add more time zones as needed
}

# Global variables for GUI elements
timezone_var = None
duration_var = None
recipient_entry = None
second_email_entry = None
merge_var = None
text_widget = None
email_entry = None
email_entry_frame = None
owner_name_entry = None
user_email = None

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

def parse_datetime(event_time):
    if 'dateTime' in event_time:
        return datetime.fromisoformat(event_time['dateTime'])  # Offset-aware
    elif 'date' in event_time:
        atlantic = pytz.timezone('America/Halifax')
        return atlantic.localize(datetime.fromisoformat(event_time['date'] + 'T00:00:00'))
    else:
        raise ValueError("Invalid event time format")

def is_ignored_event(event):
    ignored_titles = ["Office", "Home"]
    if 'summary' in event and event['summary'] in ignored_titles:
        return True
    return False

def next_15_minute_increment(dt):
    if dt.minute % 15 == 0 and dt.second == 0 and dt.microsecond == 0:
        return dt
    else:
        minutes = (dt.minute // 15) * 15 + 15
        return dt.replace(minute=0, second=0, microsecond=0) + timedelta(minutes=minutes)

def previous_15_minute_increment(dt):
    minutes = (dt.minute // 15) * 15
    return dt.replace(minute=0, second=0, microsecond=0) + timedelta(minutes=minutes)

def get_open_slots(events, day_start, day_end):
    open_slots = []
    current_start = day_start
    events.sort(key=lambda e: parse_datetime(e['start']))

    for event in events:
        if is_ignored_event(event):
            continue
        start = parse_datetime(event['start'])
        end = parse_datetime(event['end'])
        if start > current_start:
            open_slots.append((current_start, start))
        current_start = max(current_start, end)

    if current_start < day_end:
        open_slots.append((current_start, day_end))

    return open_slots

def get_events_from_calendar(calendar_id, time_min_iso, time_max_iso, service):
    events_result = service.events().list(
        calendarId=calendar_id,
        timeMin=time_min_iso,
        timeMax=time_max_iso,
        singleEvents=True,
        orderBy='startTime').execute()
    return events_result.get('items', [])

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

def get_availability(calendar_id, week_offset=0, timezone_name='Atlantic Standard Time', duration_minutes=30):
    service = build_service()
    atlantic = pytz.timezone('America/Halifax')
    target_timezone = pytz.timezone(TIMEZONES.get(timezone_name, 'America/Halifax'))
    today = datetime.now(atlantic)
    monday = today - timedelta(days=today.weekday()) + timedelta(weeks=week_offset)
    all_slots = []

    for day_offset in range(5):  # Monday to Friday
        current_day = monday + timedelta(days=day_offset)
        time_min = datetime(current_day.year, current_day.month, current_day.day, 11, 0, tzinfo=atlantic)
        time_max = datetime(current_day.year, current_day.month, current_day.day, 17, 0, tzinfo=atlantic)

        time_min_iso = time_min.isoformat()
        time_max_iso = time_max.isoformat()

        events = get_events_from_calendar(calendar_id, time_min_iso, time_max_iso, service)
        open_slots = get_open_slots(events, time_min, time_max)

        for slot in open_slots:
            slot_start = slot[0]
            slot_end = slot[1]

            aligned_start = next_15_minute_increment(slot_start)
            aligned_end = previous_15_minute_increment(slot_end)

            t = aligned_start
            while t + timedelta(minutes=duration_minutes) <= slot_end:
                all_slots.append((t, t + timedelta(minutes=duration_minutes)))
                t += timedelta(minutes=15)

    if len(all_slots) >= 5:
        selected_slots = random.sample(all_slots, 5)
    else:
        selected_slots = all_slots

    selected_slots.sort()

    availability = []
    for slot in selected_slots:
        slot_start_in_tz = slot[0].astimezone(target_timezone)
        slot_end_in_tz = slot[1].astimezone(target_timezone)
        day_str = slot_start_in_tz.strftime('%A, %B %d, %Y')
        time_str = f"{slot_start_in_tz.strftime('%I:%M %p')} - {slot_end_in_tz.strftime('%I:%M %p')} {timezone_name}"
        availability.append(f"{day_str}:\n{time_str}")

    return availability

def find_common_slots(slots1, slots2):
    common_slots = []
    i, j = 0, 0

    while i < len(slots1) and j < len(slots2):
        start1, end1 = slots1[i]
        start2, end2 = slots2[j]

        latest_start = max(start1, start2)
        earliest_end = min(end1, end2)

        if latest_start < earliest_end:
            common_slots.append((latest_start, earliest_end))

        if end1 < end2:
            i += 1
        else:
            j += 1

    return common_slots

def get_common_free_slots(calendar_id1, calendar_id2, week_offset=0, timezone_name='Atlantic Standard Time', duration_minutes=30):
    service = build_service()
    atlantic = pytz.timezone('America/Halifax')
    target_timezone = pytz.timezone(TIMEZONES.get(timezone_name, 'America/Halifax'))
    today = datetime.now(atlantic)
    monday = today - timedelta(days=today.weekday()) + timedelta(weeks=week_offset)
    all_slots = []

    for day_offset in range(5):
        current_day = monday + timedelta(days=day_offset)
        time_min = datetime(current_day.year, current_day.month, current_day.day, 11, 0, tzinfo=atlantic)
        time_max = datetime(current_day.year, current_day.month, current_day.day, 17, 0, tzinfo=atlantic)

        time_min_iso = time_min.isoformat()
        time_max_iso = time_max.isoformat()

        events_person1 = get_events_from_calendar(calendar_id1, time_min_iso, time_max_iso, service)
        events_person2 = get_events_from_calendar(calendar_id2, time_min_iso, time_max_iso, service)

        open_slots_person1 = get_open_slots(events_person1, time_min, time_max)
        open_slots_person2 = get_open_slots(events_person2, time_min, time_max)

        common_slots = find_common_slots(open_slots_person1, open_slots_person2)

        for slot in common_slots:
            slot_start = slot[0]
            slot_end = slot[1]

            aligned_start = next_15_minute_increment(slot_start)
            aligned_end = previous_15_minute_increment(slot_end)

            t = aligned_start
            while t + timedelta(minutes=duration_minutes) <= slot_end:
                all_slots.append((t, t + timedelta(minutes=duration_minutes)))
                t += timedelta(minutes=15)

    if len(all_slots) >= 5:
        selected_slots = random.sample(all_slots, 5)
    else:
        selected_slots = all_slots

    selected_slots.sort()

    availability = []
    for slot in selected_slots:
        slot_start_in_tz = slot[0].astimezone(target_timezone)
        slot_end_in_tz = slot[1].astimezone(target_timezone)
        day_str = slot_start_in_tz.strftime('%A, %B %d, %Y')
        time_str = f"{slot_start_in_tz.strftime('%I:%M %p')} - {slot_end_in_tz.strftime('%I:%M %p')} {timezone_name}"
        availability.append(f"{day_str}:\n{time_str}")

    return availability

def show_availability(week_offset=0):
    try:
        selected_timezone = timezone_var.get()
        selected_duration_str = duration_var.get()
        selected_duration = 30 if "30" in selected_duration_str else 60

        # Get the recipient name and owner name from the main window
        recipient_name = recipient_entry.get().strip()
        if not recipient_name:
            recipient_name = "there"
        
        owner_name = owner_name_entry.get().strip()
        if not owner_name:
            owner_name = "my"  # Default to "my availability" if none provided

        merge = (merge_var.get() == 1)

        if merge:
            # If merging, we need the second email
            calendar_id2 = second_email_entry.get().strip()
            if not calendar_id2:
                messagebox.showwarning("Input Required", "Please enter the second person's email address for merged availability.")
                return
            availability = get_common_free_slots(user_email, calendar_id2, week_offset, selected_timezone, selected_duration)
        else:
            availability = get_availability(user_email, week_offset, selected_timezone, selected_duration)

        period_str = "this week" if week_offset == 0 else "next week"

        if merge:
            greeting_line = f"Hi {recipient_name}, here is our availability for {period_str}:\n\n"
        else:
            # Use the owner_name if provided, otherwise 'my'
            greeting_line = f"Hi {recipient_name}, here is {owner_name}'s availability for {period_str}:\n\n"

        availability_text = "\n\n".join(availability)

        text_widget.delete(1.0, tk.END)
        text_widget.insert(tk.END, greeting_line + availability_text)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def copy_to_clipboard():
    availability_text = text_widget.get(1.0, tk.END)
    pyperclip.copy(availability_text)
    messagebox.showinfo("Copied", "Availability copied to clipboard.")

def initialize_app():
    global user_email
    user_email = email_entry.get().strip()
    if not user_email:
        messagebox.showwarning("Input Required", "Please enter your email address.")
        return
    email_entry_frame.destroy()
    display_main_gui()

def display_main_gui():
    global timezone_var, duration_var, recipient_entry, second_email_entry, merge_var, text_widget, owner_name_entry

    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    header_label = ttk.Label(main_frame, text="Google Calendar Availability", font=('Arial', 16, 'bold'))
    header_label.pack(pady=(0, 10))

    # Owner name field
    owner_name_label = ttk.Label(main_frame, text="Owner's Name (leave blank for 'my availability'):")
    owner_name_label.pack(pady=5)
    owner_name_entry = ttk.Entry(main_frame, width=40)
    owner_name_entry.pack(pady=5)

    # Recipient name field
    recipient_label = ttk.Label(main_frame, text="Recipient's Name:")
    recipient_label.pack(pady=5)
    recipient_entry = ttk.Entry(main_frame, width=40)
    recipient_entry.pack(pady=5)

    # Second person's email field (for merging)
    second_email_label = ttk.Label(main_frame, text="Second Person's Email (for merge):")
    second_email_label.pack(pady=5)
    second_email_entry = ttk.Entry(main_frame, width=40)
    second_email_entry.pack(pady=5)

    # Checkbox for merge
    merge_var = tk.IntVar(value=0)
    merge_checkbox = ttk.Checkbutton(main_frame, text="Merge Availability", variable=merge_var)
    merge_checkbox.pack(pady=5)

    # Time zone selection
    timezone_label = ttk.Label(main_frame, text="Select Time Zone:")
    timezone_label.pack(pady=5)
    timezone_var = tk.StringVar()
    timezone_var.set("Atlantic Standard Time")
    timezone_combobox = ttk.Combobox(main_frame, textvariable=timezone_var, values=list(TIMEZONES.keys()), state='readonly')
    timezone_combobox.pack(pady=5)

    # Meeting duration selection
    duration_label = ttk.Label(main_frame, text="Select Meeting Duration:")
    duration_label.pack(pady=5)
    duration_var = tk.StringVar()
    duration_var.set("30 minutes")
    duration_options = ["30 minutes", "1 hour"]
    duration_combobox = ttk.Combobox(main_frame, textvariable=duration_var, values=duration_options, state='readonly')
    duration_combobox.pack(pady=5)

    buttons_frame = ttk.Frame(main_frame)
    buttons_frame.pack(pady=10)

    fetch_this_week_button = ttk.Button(buttons_frame, text="This Week's Availability",
                                        command=lambda: show_availability(week_offset=0))
    fetch_next_week_button = ttk.Button(buttons_frame, text="Next Week's Availability",
                                        command=lambda: show_availability(week_offset=1))
    fetch_this_week_button.grid(row=0, column=0, padx=5, pady=5)
    fetch_next_week_button.grid(row=0, column=1, padx=5, pady=5)

    copy_button = ttk.Button(main_frame, text="Copy to Clipboard", command=copy_to_clipboard)
    copy_button.pack(pady=10)

    text_frame = ttk.Frame(main_frame)
    text_frame.pack(fill=tk.BOTH, expand=True)

    text_widget_scrollbar = ttk.Scrollbar(text_frame)
    text_widget_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    text_widget = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=text_widget_scrollbar.set)
    text_widget.pack(fill=tk.BOTH, expand=True)
    text_widget_scrollbar.config(command=text_widget.yview)

    # Store references globally
    globals()['recipient_entry'] = recipient_entry
    globals()['second_email_entry'] = second_email_entry
    globals()['merge_var'] = merge_var
    globals()['owner_name_entry'] = owner_name_entry

def create_gui():
    global root, email_entry_frame, email_entry
    root = tk.Tk()
    root.title("Google Calendar Availability")

    window_width = 600
    window_height = 700
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

    style = ttk.Style()
    style.theme_use('clam')

    bg_image_path = get_resource_path("wallpaper.png")
    bg_image = Image.open(bg_image_path)
    bg_image = bg_image.resize((window_width, window_height), Image.Resampling.LANCZOS)
    bg_image_tk = ImageTk.PhotoImage(bg_image)

    background_label = tk.Label(root, image=bg_image_tk)
    background_label.place(relwidth=1, relheight=1)
    background_label.image = bg_image_tk  # Keep reference

    overlay_frame = tk.Frame(root, bg='white', bd=0, highlightthickness=0)
    overlay_frame.place(relx=0.5, rely=0.5, anchor='center')

    logo_path = get_resource_path("logo.png")
    logo_image = Image.open(logo_path)
    logo_image = logo_image.resize((150, 150), Image.Resampling.LANCZOS)
    logo_image_tk = ImageTk.PhotoImage(logo_image)

    logo_label = ttk.Label(overlay_frame, image=logo_image_tk)
    logo_label.pack(pady=10)
    logo_label.image = logo_image_tk  # Keep reference

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
