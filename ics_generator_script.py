import pandas as pd
import os
from icalendar import Calendar, Event, Alarm
from datetime import timedelta
from dotenv import load_dotenv

calendar = Calendar()
calendar.add('version', '2.0')

def is_excel_file_valid(file_path):
    if not file_path:
        print("❗ Unable to find the source file. Please ensure the EXCEL_FILE_PATH is specified in the .env file.")
        return False
    
    if not file_path.endswith(".xlsx"):
        print("❗ The specified file is not an Excel file. Please provide a valid .xlsx file.")
        return False

    if not os.path.isfile(file_path):
        print(f"❗ The file '{file_path}' does not exist. Please check the path and try again.")
        return False
    
    return True

# Get the excel file path from .env
load_dotenv()
path_to_excel_file = os.getenv('EXCEL_FILE_PATH')
if (not is_excel_file_valid(path_to_excel_file)):
    exit(1)

data_frame = pd.read_excel(path_to_excel_file, index_col=None, sheet_name="Due Dates")
data_frame['Due Date'] = pd.to_datetime(data_frame['Due Date'], errors='coerce')
print("\n::STARTED::")

for index, row in data_frame.iterrows():
    try:
        deadline_datetime = row['Due Date']
        assignment = row['Assignment']
        course = row['Course']
    except KeyError:
        print("❗ There was an error reading the excel file...")
        exit(1)
    
    # If a due date is empty, skip adding that event
    if pd.isna(deadline_datetime):
        print(f"❌ Event not created for {assignment} - {course} (Unknown Due Date)")
        continue

    event = Event()
    event.add('summary', f"{assignment} {course}")
    event.add('dtstart', deadline_datetime.date())
    event.add('dtend', deadline_datetime.date())
    event.add('description', f"Assignment: {assignment}\nCourse: {course}")
    event.add('status', 'CONFIRMED')
    event.add('transp', 'TRANSPARENT')
    event.add('class', 'PRIVATE')

    # Adding reminders
    reminder_offsets = [0, 1, 2, 4, 6]  # Days before the event
    for offset in reminder_offsets:
        
        # Calculating a reminder time of 9:00 AM
        reminder_time = deadline_datetime - timedelta(days=offset)
        reminder_time = reminder_time.replace(hour=9, minute=0, second=0)

        alarm = Alarm()
        alarm.add('trigger', reminder_time - deadline_datetime)
        alarm.add('action', 'DISPLAY')
        alarm.add('description', f"Reminder for {assignment} {course}")
        event.add_component(alarm)
    
    calendar.add_component(event)

# Exporting to a '.ics' file
ics_file = 'assignments_calendar.ics'
with open(ics_file, 'wb') as f:
    f.write(calendar.to_ical())

print(f"✔️  Calendar file created: '{ics_file}'")
print("::FINISHED::\n")
