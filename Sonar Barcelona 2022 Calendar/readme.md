# Sonar 2022 program in Google Calendar / Outlook format
[Sonar](https://sonar.es/) is one of the world's largest electronic music festivals. It is held in Barcelona, Spain every summer. Over 3 days and 3 nights, there are [150+ performances of artists, DJs and speakers](https://sonar.es/en/2022/artists) spread over [several locations and stages](https://sonar.es/en/2022/showcases).

Sonar 2022 takes place on June 16-18 2022.

If you are interested to attend specific events, you need to have the event information in a manageable form. One way is of course to refer to the [Sonar 2022 timetable on the festival's website](https://sonar.es/en/2022/schedules). Another way would be to pick a list of the events you like and save them to a calendar. Then you will be able to see what is happening when and where.

Adding hundreds of events to your calendar by hand is quite a mouthful. You need a way to import the events automatically. And this is a perfect use case for web scraping, which I do with powershell in quick and dirty but working way.

## Prerequisites
I use Windows and PowerShell to do the web scraping.

## How it works

1. Download all webpages with artist and schedule information
2. Extract this schedule information into a structured form
3. Transform the structured schedule information into a format that can be imported to a calendar (an Outlook CSV)
4. [manual] Import it to a public gcal and share the calendar with your friends

## Import to calendar
I find Google Calendar (gcal) a perfect way to create shared public calendars.

And gcal supports the [import of structured files](https://calendar.google.com/calendar/u/0/r/settings/export), too! One of the easiest formats for me, being on Windows, is Outlook CSV format, and gcal takes it in out of the box! So I wrote a script that collects (scraps) the information from Sonar website and creates a CSV file to import to gcal.

## Does this script work with 100% of Sonar pages?
No, but it is still good enough.
* I noted that [the page of Santos Bacana](https://sonar.es/en/2022/artists/santos-bacana) does not contain time information and therefore cannot be imported.
* Some parsed events are duplicated. This is not a huge problem, however, because gcal takes care of the duplicates.
* Some other event pages, e.g. [Museum of Sound listening experience](https://sonar.es/en/2022/artists/museum-of-sound-listening-experience-by-mika-vainio) contain multiple time entries, and these are not taken into account.

## How the result looks
### Calendar
![how the output looks in gcal](https://i.imgur.com/O7OpKIO.png)
### Details of single event
![how event details look in gcal](https://i.imgur.com/jUycrfr.png)

## What else can be done
* Schedule information actually can be taken straight from the Sonar timetable page. I did not go for that because the data structure in that page is `Date | Day/Night -> Venue -> List of events`, and I would have to figure out how to attribute child elements to the parent. Whereas in individual artist pages complete event information is available.
* Separate calendars can be created per Venue, Day/Night, etc. I actually initially created a calendar out of our own carefully curated selection of events. In any case, after the event information is collected in a powershell object, it can be sliced any way possible.
* Times can be added with time zone information. I do not do that because I am in the same time zone.
