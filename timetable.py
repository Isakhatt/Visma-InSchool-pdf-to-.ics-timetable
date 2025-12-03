import sys
import pytz
from copy import deepcopy
import pymupdf
import datetime
from uuid import uuid4
import tkinter as tk
from tkinter import filedialog, ttk
from tzlocal import get_localzone

class VeventBlock:
    def __init__(self, start_time: str, end_time: str, location: str, subject: str):
        """
        Constructor

        start_time format: HH:MM
        end_time format: HH:MM
        location: string
        subject: string
        """
        self.start_time = start_time
        self.end_time = end_time
        self.location = location
        self.subject = subject
    
    def showInfo(self):
        """
        Prints general info about the instanced object
        """
        print(f"Name: {self.subject} Starts at: {self.start_time} Ends at: {self.end_time} Location: {self.location}")


    def startMinutesPastMidnight(self) -> int:
        """
        Returns the minutes that have passed, when this event starts, after midnight
        """
        return int(self.start_time[0] + self.start_time[1]) * 60 + int(self.start_time[3] + self.start_time[4])
    
    def endMinutesPastMidnight(self) -> int:
        """
        Returns the minutes that have passed, when this event starts, after midnight
        """
        return int(self.end_time[0] + self.end_time[1]) * 60 + int(self.end_time[3] + self.end_time[4])



def timeconvert(visma_time: str, visma_date: str) -> str:
    """
    Converts Visma time and date format to iCalendar datetime format.
    Input time format: HH:MM
    Input date format: DD.MM.YYYY
    Output format: YYYYMMDDTHHMMSS
    """
    date = datetime.datetime.strptime(visma_date, r'%d.%m.%Y')  # Handles 4-digit year
    time = datetime.datetime.strptime(visma_time, r'%H:%M').time()
    fulltime = datetime.datetime.combine(date, time)
    return datetime.datetime.strftime(fulltime, r'%Y%m%dT%H%M%S')



# Initialize the tk window
root = tk.Tk()
root.geometry("250x150")
root.title("Select Timezone")

# Define the timezone combobox
cb = ttk.Combobox(root, values=pytz.common_timezones, width=30)
cb.set(str(get_localzone()))  # default value
cb.pack(pady=10)

# Label for status messages
lbl = tk.Label(root, text="Please select your timezone.")
lbl.pack(pady=5)

# Variable to store the selected timezone
TIMEZONE = None

def confirm_timezone():
    """Save the timezone and close the window."""
    global TIMEZONE
    TIMEZONE = cb.get()
    print(f"Selected timezone: {TIMEZONE}")  # or handle it elsewhere
    root.destroy()  # closes the window and allows program to continue

# Confirm button
tk.Button(root, text="Confirm", command=confirm_timezone).pack(pady=10)

# Start the GUI loop
root.mainloop()

# Program continues after window is closed
print(f"Using timezone: {TIMEZONE}")


# Define the input PDF file path
pdf_path = filedialog.askopenfilename(title='Select the downloaded timetable from Visma InSchool', filetypes={("pdf", ".pdf")})


# Read the PDF and extract text. Stop if no pdf is selected
try:
    doc = pymupdf.open(pdf_path)
    pdf_text = "".join(doc[0].get_text())
except Exception as e:
    sys.exit(f"Exiting: {e}")


lines = pdf_text.splitlines() # Split the text into lines for processing
# Each lesson is assumed to span 4 lines: time | location, subject, code, type

# Remove the days of the week on the top of the document
for a in range(5):
    lines.pop(0)

download_date = datetime.datetime.strptime(lines[-1][-10:], '%d.%m.%Y') # Define download date (it is shown in Visma)

lines.pop(-1) # Remove the last line

# Find difference between the download date and the one the first lesson is on.
current_date = download_date - datetime.timedelta(days=download_date.weekday()) 

# Prepare timestamp for DTSTAMP field (convert to right format)
timestamp = datetime.datetime.strftime(datetime.datetime.now(datetime.timezone.utc), r'%Y%m%dT%H%M%SZ') 

# Initialize the iCalendar content
ical = 'BEGIN:VCALENDAR\r\nPRODID:-//Visma Inschool timetable to iCalendar//EN\r\nVERSION:2.0\r\n' 

# Define values
prev_prev_event = None
prev_event = VeventBlock("00:00", "23:59", "", "")
next_event = prev_event

# Main loop
a = 0
while a < len(lines) - 1:
    # Check if the first four letters are in time format; two numbers, colon, two numbers
    try: 
        if type(int(lines[a][0] + lines[a][1] + lines[a][3] + lines[a][4])) == int and lines[a][2] == ":":
            # Define location, start_time, end_time and subject
            time_range, location = lines[a].split("|")
            start_time, end_time = [t.strip() for t in time_range.split("-")]
            subject = lines[a + 1].strip()
            
            # Instancing current event
            current_event = VeventBlock(start_time, end_time, location, subject)
            
            # Check for duplicate event (after combining events)
            if current_event.endMinutesPastMidnight() == prev_event.endMinutesPastMidnight() and current_event.subject == prev_event.subject:
                a += 1
            else: 

                # Find next event block
                b = a + 2
                while b < len(lines) - 1:
                    try:
                        if type(int(lines[b][0] + lines[b][1] + lines[b][3] + lines[b][4])) == int and lines[b][2] == ":":
                            time_range, location = lines[b].split("|")
                            start_time, end_time = [t.strip() for t in time_range.split("-")]
                            subject = lines[b + 1].strip()

                            next_event = VeventBlock(start_time, end_time, location, subject)
                            b = len(lines)
                        
                        else:
                            b += 1
                    except Exception:
                        b += 1


                # Check if new day, if this start is less than previous. Defines start_times as integers representing the passed minutes since 00:00
                if current_event.startMinutesPastMidnight() < prev_event.startMinutesPastMidnight():
                    current_date += datetime.timedelta(days=1)


                # Check if next event is a continuation of the previous and has the same name
                if current_event.end_time == next_event.start_time and current_event.subject == next_event.subject:
                    # Extend the current event to the end of the next
                    current_event.end_time = next_event.end_time

                # Defines previous event as this one (this event will be written over)
                if prev_prev_event == None:
                    prev_prev_event = deepcopy(current_event)
                    prev_prev_event.start_time = "00:00"
                    prev_prev_event.end_time = "00:00"
                else:
                    prev_prev_event = deepcopy(prev_event)
                prev_event = deepcopy(current_event)
                

                # Add 45 minutes if the class is 45 minutes long and the last of the day or the last day of the timetable. This is in case the pdf is longer than 1 page and follows the class system of Norwegian high schools
                # Seeing if there is not a next event block in the list first (in case there is an error reading the next event of the block that does not exist)
                if len(lines) - a < 5:
                    current_event.end_time = (datetime.datetime.combine(datetime.datetime.today(), datetime.datetime.strptime(current_event.end_time, '%H:%M').time()) + datetime.timedelta(minutes=45)).time().strftime('%H:%M')
                elif next_event.start_time < current_event.start_time and current_event.endMinutesPastMidnight() - current_event.startMinutesPastMidnight() < 90:
                    current_event.end_time = (datetime.datetime.combine(datetime.datetime.today(), datetime.datetime.strptime(current_event.end_time, '%H:%M').time()) + datetime.timedelta(minutes=45)).time().strftime('%H:%M')

                # Defines date string in iCalendar format
                date_str = current_date.strftime("%d.%m.%Y")


                # Show info about event
                # current_event.showInfo()

                # Generate vevent block. As long as the current event starts after the previous event ended, the program adds an alarm. 
                current_event.showInfo()
                if current_event.start_time != prev_prev_event.end_time and current_event.start_time != prev_prev_event.start_time: 
                    vevent = [
                        "BEGIN:VEVENT",
                        f"SUMMARY:{current_event.subject}",
                        f"DTSTART;TZID={TIMEZONE}:{timeconvert(current_event.start_time, date_str)}",
                        f"DTEND;TZID={TIMEZONE}:{timeconvert(current_event.end_time, date_str)}",
                        f"LOCATION:{location.strip()}",
                        f"UID:fromvisma_{uuid4()}",
                        f"DTSTAMP:{timestamp}",
                        "BEGIN:VALARM",
                        f"TRIGGER:-PT15M", # Alarm starts 15 min before
                        "ACTION:DISPLAY",
                        "DESCRIPTION:Reminder",
                        "END:VALARM",
                        "END:VEVENT\r\n"
                    ]
                else:
                    vevent = [
                        "BEGIN:VEVENT",
                        f"SUMMARY:{current_event.subject}",
                        f"DTSTART;TZID={TIMEZONE}:{timeconvert(current_event.start_time, date_str)}",
                        f"DTEND;TZID={TIMEZONE}:{timeconvert(current_event.end_time, date_str)}",
                        f"LOCATION:{location.strip()}",
                        f"UID:fromvisma_{uuid4()}",
                        f"DTSTAMP:{timestamp}",
                        "END:VEVENT\r\n"
                    ]
                ical += "\r\n".join(vevent) # Adds the block to the file
                a += 1
        
        else:
            # If lines[a] isn't the correct time format, but still passes the if test
            a += 1
    except ValueError as e:
        # Expected, happens on non-time lines
        a += 1
        continue
    except Exception as e:
        # Unexpected errors
        print("Error: {e}")
        a += 1

print("done")

# Finalize the calendar
ical += 'END:VCALENDAR\r\n'

# Save the iCalendar content to a file
output_filename = filedialog.asksaveasfilename(
    title="Save the .ics file as", 
    defaultextension=".ics", 
    filetypes=[("iCalendar files", "*.ics")]
)
with open(output_filename, "w") as f:
    f.write(ical)
