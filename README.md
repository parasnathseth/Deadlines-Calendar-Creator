# Deadline Calendar Creator

## **What is this?**
This is a Python script that generates an `.ics` calendar file from an Excel spreadsheet containing assignment due dates. The `.ics` file can then be imported into calendar applications.

## **Why did I make this?**

I absolutely love using Excel and Calendar apps to stay on top of my deadlines and events.

At the beginnning of every semester of university, I put together a spreadsheet containing all my course-related deadlines and dates for major events. I then end up spending a *really* long time creating events for each deadline.

I created this simple script to automate the manual process of adding assignment due dates to my calendar, saving me some time each semester!

## **What does this do?**
When executed, this script performs the following steps:
1. Reads an Excel spreadsheet containing due dates, assignment titles, and course names.
2. Creates an `.ics` file containing calendar events based on the information from the spreadsheet.
3. Adds reminders for each event at the following intervals:
   - On the day of the event (9:00 AM).
   - 1 day before (9:00 AM).
   - 2 days before (9:00 AM).
   - 4 days before (9:00 AM).
   - 6 days before (9:00 AM).
4. Sets the events as **all-day events** with **private visibility**.
5. Logs any rows with missing due dates to the console and skips them.

## **Assumptions**

### **1. Excel File Format**
The script assumes the Excel spreadsheet has the following structure, with 3-columns:

| **Column Name (case-sensitive)** | **Description**                                           |
|------------------|-----------------------------------------------------------|
| `Due Date`       | The due date of the assignment (e.g., `Thursday, January 16, 2025`). |
| `Assignment`     | The title or name of the assignment.                      |
| `Course`         | The course name or identifier.                           |

- Dates must follow the format: "[Day of the Week], [Month] [Day], [Year]" (e.g., `Thursday, January 16, 2025`).
- The column names must match **exactly** (`Due Date`, `Assignment`, `Course`).

### **2. Missing Data**
- If the `Due Date` is missing for a row, the script will:
  - Skip creating an event for that row.
  - Print a message to the console indicating the skipped event

### **3. Event Settings**
- All events are created as **all-day events**.
- Events are marked as **private**, meaning they won’t be shared with others viewing your calendar.
- Events do not repeat—they are standalone deadlines.

### **4. Reminders**
Each event includes the following reminders:
- On the day of the event at **9:00 AM**.
- 1, 2, 4, and 6 days before the event, all at **9:00 AM**.


## **How to Use**

### **1. Install Dependencies**
Make sure you have Python installed, and install the required dependencies:
```bash
pip install pandas openpyxl icalendar python-dotenv
```

### **2. Setup .env file**
Copy the provided `example.env` file to `.env` in the project directory:
```bash
   cp example.env .env
```

### **3. Run the Script**

### **4. Verify the Output**
- Make sure that important events are **not** skipped over, and that there are no error messages in the console output

### **5. Import the `.ics` File**

- The `.ics` file should now be ready for you to import into your calendar app(s)!