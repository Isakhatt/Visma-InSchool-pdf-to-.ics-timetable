import pytz
from copy import deepcopy
import pymupdf
import datetime
from uuid import uuid4
import tkinter as tk
from tkinter import filedialog, ttk
from tzlocal import get_localzone

class VeventBlock:
    def __init__(self, start_time: str, end_time: str, location: str | None, subject: str):
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

# Define the timezone box
cb = ttk.Combobox(root, values=pytz.common_timezones, width=30)
cb.set(str(get_localzone()))  # default value
cb.pack(pady=10)

# Label for status messages
lbl = tk.Label(root, text="Please select your timezone.")
lbl.pack(pady=5)

TIMEZONE = None

def confirm_timezone():
    """Save the timezone and close the window."""
    global TIMEZONE
    TIMEZONE = cb.get()
    root.destroy()  # closes the window

# Confirm button
tk.Button(root, text="Confirm", command=confirm_timezone).pack(pady=10)

# Start the GUI loop
root.mainloop()

# Program continues after window is closed
print(f"Using timezone: {TIMEZONE}")


# Define the input PDF file path
pdf_path = filedialog.askopenfilename(title='Select the downloaded timetable from Visma InSchool', filetypes={("pdf", ".pdf")})

# Read the PDF and extract text. Stop if no pdf is selected
doc = pymupdf.open(pdf_path)

lines = []


for i in range(len(doc)):
    page: pymupdf.Page = doc[i]

    # Define columns to read text from, each representing their own weekday
    page_rect = page.rect
    width = page_rect.width
    height = page_rect.height
    rectangles = []

    y = 50 # Distance from top to remove

    rectangles.append(pymupdf.Rect(0, y, round(width/5), height))
    rectangles.append(pymupdf.Rect(round(width/5), y, round(width/5)*2, height))
    rectangles.append(pymupdf.Rect(round(width/5)*2, y, round(width/5)*3, height))
    rectangles.append(pymupdf.Rect(round(width/5)*3, y, round(width/5)*4, height))
    rectangles.append(pymupdf.Rect(round(width/5)*4, y, round(width/5)*5, height))

    for rect in rectangles:
        text = page.get_textbox(rect)
        lines.append(text.splitlines()) # Split the text into lines for processing
        # Each lesson is assumed to span 4 lines: time | location, subject, code, type


# Define first date in timetable
page: pymupdf.Page = doc[0]
page_rect = page.rect
page_rect.y1 = 20
header = page.get_textbox(page_rect)
header = header.split(" - ")

start_date = datetime.datetime.strptime(header[-1], r'%d.%m.%Y') - datetime.timedelta(days=datetime.datetime.strptime(header[-1], r'%d.%m.%Y').weekday())

# Prepare timestamp for DTSTAMP field (convert to right format)
timestamp = datetime.datetime.strftime(datetime.datetime.now(datetime.timezone.utc), r'%Y%m%dT%H%M%SZ') 

# Initialize the iCalendar content
ical = 'BEGIN:VCALENDAR\r\nPRODID:-//Visma Inschool timetable to iCalendar//EN\r\nVERSION:2.0\r\n' 

# Define values
prev_event = VeventBlock("00:00", "23:59", "", "")
next_event = prev_event

# Main loop
for i in range(5):
    # Add days to start_date
    current_date = start_date + datetime.timedelta(days=i)

    a = 0
    while a < len(lines[i]) - 1:
        # Check if the first four letters are in time format; two numbers, colon, two numbers
        try: 
            if type(int(lines[i][a][0] + lines[i][a][1] + lines[i][a][3] + lines[i][a][4])) == int and lines[i][a][2] == ":":
                # Define location, start_time, end_time and subject
                if "|" in lines[i][a]:
                    time_range, location = lines[i][a].split("|")
                    location = location.strip()
                else:
                    time_range = lines[i][a]
                    location = None
                
                start_time, end_time = [t.strip() for t in time_range.split("-")]
                subject = lines[i][a + 1].strip()
                
                # Instancing current event
                current_event = VeventBlock(start_time, end_time, location, subject)
                
                # Check for duplicate event (after combining events)
                if current_event.endMinutesPastMidnight() == prev_event.endMinutesPastMidnight() and current_event.subject == prev_event.subject:
                    a += 1
                else: 
    
                    # Find next event block
                    b = a + 2
                    while b < len(lines[i]) - 1:
                        try:
                            if type(int(lines[i][b][0] + lines[i][b][1] + lines[i][b][3] + lines[i][b][4])) == int and lines[i][b][2] == ":":
                                if "|" in lines[i][b]:
                                    time_range, location = lines[i][b].split("|")
                                    location = location.strip()
                                else:
                                    time_range = lines[i][b]
                                    location = None
                                start_time, end_time = [t.strip() for t in time_range.split("-")]
                                subject = lines[i][b + 1].strip()
    
                                next_event = VeventBlock(start_time, end_time, location, subject)
                                b = len(lines[i])
                            
                            else:
                                b += 1
                        except Exception:
                            b += 1

    
                    # Check if next event is a continuation of the previous and has the same name
                    if current_event.end_time == next_event.start_time and current_event.subject == next_event.subject:
                        # Extend the current event to the end of the next
                        current_event.end_time = next_event.end_time
                    
    
                    # Add 45 minutes if the class is 45 minutes long and the last of the day or the last day of the timetable. This is in case the pdf is longer than 1 page and follows the class system of Norwegian high schools
                    # Seeing if there is not a next event block in the list first (in case there is an error reading the next event of the block that does not exist)
                    if len(lines[i]) - a < 5:
                        current_event.end_time = (datetime.datetime.combine(datetime.datetime.today(), datetime.datetime.strptime(current_event.end_time, '%H:%M').time()) + datetime.timedelta(minutes=45)).time().strftime('%H:%M')
                    elif next_event.start_time < current_event.start_time and current_event.endMinutesPastMidnight() - current_event.startMinutesPastMidnight() < 90:
                        current_event.end_time = (datetime.datetime.combine(datetime.datetime.today(), datetime.datetime.strptime(current_event.end_time, '%H:%M').time()) + datetime.timedelta(minutes=45)).time().strftime('%H:%M')
    
                    # Defines date string in iCalendar format
                    date_str = current_date.strftime("%d.%m.%Y")
    
    
                    # Show info about event
                    # current_event.showInfo()
    
                    # Generate vevent block. As long as the current event starts after the previous event ended, the program adds an alarm. 
                    current_event.showInfo()
                    if current_event.start_time != prev_event.end_time and current_event.start_time != prev_event.start_time: 
                        if type(current_event.location) == str:
                            vevent = [
                                "BEGIN:VEVENT",
                                f"SUMMARY:{current_event.subject}",
                                f"DTSTART;TZID={TIMEZONE}:{timeconvert(current_event.start_time, date_str)}",
                                f"DTEND;TZID={TIMEZONE}:{timeconvert(current_event.end_time, date_str)}",
                                f"LOCATION:{current_event.location}",
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
                        if type(current_event.location) == str:
                            vevent = [
                                "BEGIN:VEVENT",
                                f"SUMMARY:{current_event.subject}",
                                f"DTSTART;TZID={TIMEZONE}:{timeconvert(current_event.start_time, date_str)}",
                                f"DTEND;TZID={TIMEZONE}:{timeconvert(current_event.end_time, date_str)}",
                                f"LOCATION:{current_event.location}",
                                f"UID:fromvisma_{uuid4()}",
                                f"DTSTAMP:{timestamp}",
                                "END:VEVENT\r\n"
                            ]
                        else:
                            vevent = [
                                "BEGIN:VEVENT",
                                f"SUMMARY:{current_event.subject}",
                                f"DTSTART;TZID={TIMEZONE}:{timeconvert(current_event.start_time, date_str)}",
                                f"DTEND;TZID={TIMEZONE}:{timeconvert(current_event.end_time, date_str)}",
                                f"UID:fromvisma_{uuid4()}",
                                f"DTSTAMP:{timestamp}",
                                "END:VEVENT\r\n"
                            ]
                    ical += "\r\n".join(vevent) # Adds the block to the file
    
                    # Check if next event is a continuation of this one, if true then skip next event
                    if next_event.end_time == current_event.end_time and next_event.subject == current_event.subject:
                        a += 4
    
                    # Define previous event
                    prev_event = deepcopy(current_event)
    
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
