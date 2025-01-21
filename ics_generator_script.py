import pandas as pd
from icalendar import Calendar, Event, Alarm
from datetime import datetime, timedelta

calendar = Calendar()
calendar.add('version', '2.0')

path_to_excel_file = ""
data_frame = pd.read_excel(path_to_excel_file, index_col=None, sheet_name="Due Dates")
data_frame['Due Date'] = pd.to_datetime(data_frame['Due Date'], errors='coerce')

for index, row in data_frame.iterrows():
    try:
        due_date = row['Due Date']
        assignment = row['Assignment']
        course = row['Course']
    except KeyError:
        print("There was an error reading the excel file...")
        exit(1)
    
    # If a due date is empty, skip adding that event
    if pd.isna(due_date):
        print(f"Event not created for {assignment} - {course}")
        continue

    event = Event()
    event.add('summary', f"{assignment} {course}")
    event.add('dtstart', due_date.date())
    event.add('dtend', due_date.date())
    event.add('description', f"Assignment: {assignment}\nCourse: {course}")
    event.add('status', 'CONFIRMED')
    event.add('class', 'PRIVATE')

    # Adding reminders
    reminder_offsets = [0, 1, 2, 4, 6]  # Days before the event
    for offset in reminder_offsets:
        
        # Calculating a reminder time of 9:00 AM
        reminder_time = due_date - timedelta(days=offset)
        reminder_time = reminder_time.replace(hour=9, minute=0, second=0)
        
        alarm = Alarm()
        alarm.add('trigger', reminder_time - due_date)
        alarm.add('action', 'DISPLAY')
        alarm.add('description', f"Reminder for {assignment} {course}")
        event.add_component(alarm)
    
    calendar.add_component(event)

# Exporting to a '.ics' file
ics_file = 'assignments_calendar.ics'
with open(ics_file, 'wb') as f:
    f.write(calendar.to_ical())

print(f".ics calendar file created: {ics_file}")
