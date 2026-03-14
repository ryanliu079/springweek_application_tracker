# config.py

# --- Search settings ---
# Searches ryanliu61799@gmail.com (work/applications inbox)
SEARCH_QUERY = """
(subject:("application received") OR subject:("thank you for applying") OR
 subject:("spring week") OR subject:("summer internship") OR subject:("off-cycle") OR
 subject:("online assessment") OR subject:("video interview") OR subject:("hackerrank") OR
 subject:("offer") OR subject:("unfortunately") OR subject:("application unsuccessful") OR
 subject:("next steps") OR subject:("application update") OR subject:("insight programme") OR
 subject:("early insight") OR subject:("registration confirmed") OR subject:("event application") OR
 subject:("not progress") OR subject:("not been successful") OR subject:("on this occasion") OR
 subject:("stage 1 assessment") OR subject:("discover citadel") OR subject:("future focus") OR
 from:(noreply@greenhouse.io) OR from:(noreply@lever.co) OR from:(donotreply@workday.com) OR
 from:(noreply@morganstanley.tal.net) OR from:(noreply@fidelityinternational.tal.net) OR
 from:(no-reply@drwrecruiting.com) OR from:(noreply@jefferies.tal.net) OR
 from:(noreply@campuscareers.bofa.com) OR from:(recruitment@optiver.com) OR
 from:(no-reply@optiver.com) OR from:(barclays@myworkday.com) OR
 from:(no-reply@talent-citadel.com) OR from:(ye+db@yello.co))
"""
DATE_CUTOFF = "2024/09/01"
MAX_THREADS = 300

# --- Output ---
OUTPUT_XLSX = "applications.xlsx"
SHEET_NAME = "Applications"

# --- Calendar ---
# Events are added to rliu07979@gmail.com (personal calendar)
CALENDAR_ID = "rliu07979@gmail.com"
TIMEZONE = "Europe/London"

# --- Status taxonomy ---
STATUSES = [
    "Applied",
    "Online Assessment",
    "First Round Interview",
    "Second Round Interview",
    "Assessment Centre",
    "Offer",
    "Rejected",
    "Withdrawn",
    "Pending",
]

# --- Role types to track ---
ROLE_TYPES = [
    "Spring Week",
    "Summer Internship",
    "Off-Cycle Internship",
    "Graduate Scheme",
    "Part-Time / Other",
]
