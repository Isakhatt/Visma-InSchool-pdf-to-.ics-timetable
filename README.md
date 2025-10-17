# Visma InSchool pdf to .ics timetable
Python script that makes a .ics iCalander file with the timetable in it from the Visma InSchool timetable. 

Script currently known working with google calendar. 

The script adds 45 minutes if a class is only 45 minutes and at the end because the script cannot read multiple pages and thus has to add the second half of the class. All classes are divided into 45 minute segments in Norwegian high schools as of 2025. 

The timezone has to specified manually according to the iCalendar requirements. Eks: Europe/Oslo
