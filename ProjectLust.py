import json
import csv
import sys
import smtplib
import datetime
import time
from pyluach import dates, hebrewcal, parshios
from dateutil import parser
from convertdate import hebrew
from teach_heb import heb_month, heb_day
from MASSAGE import MASSAGE
from email.mime.text import MIMEText

# !/usr/bin/env python
# -*- coding: utf-8 -*-

# Load all the settings from the json file
with open('settings.json') as s:
    obj = s.read()
    settings = json.loads(obj)
    SENDER_EMAIL = settings["SENDER_EMAIL"]
    SENDER_PASSWORD = settings["SENDER_PASSWORD"]
    TIME = settings["TIME"]

# Get today and convert it to Hebrew
today_gregorian_date = datetime.datetime.now()
today_hebrew_date = dates.HebrewDate.today()
year = hebrew.from_gregorian(today_gregorian_date.year, today_gregorian_date.month, today_gregorian_date.day)
year = year[0]  # This Hebrew Year
two_weeks_from_today = today_hebrew_date + 14

# Collect all the deceased info in separated arrays
deceased_month = []
deceased_day = []
deceased_donor = []
deceased_relation = []
deceased_name = []
deceased_SoD = []
deceased_email = []

# Array for ALL the email addresses
all_emails = []


# define a function that removes duplicates from the all_emails array
def remove_duplicates(values):
    output = []
    seen = set()
    for value in values:
        # If value has not been encountered yet,
        # ... add it to both list and set.
        if value not in seen:
            output.append(value)
            seen.add(value)
    return output


# define a function that fix the format from the csv file to the dates.HebrewDate function
def from_csv_to_date(month, day):
    return dates.HebrewDate(year, month, day)


# Self explanatory
def check_if_within_14_days_from_today(day_to_chack):
    if today_hebrew_date <= day_to_chack <= two_weeks_from_today:
        return True
    else:
        return False


# We need to open the currect file according to if this year is a leap year or not
if hebrewcal.HebrewDate._is_leap(year):
    file_name = 'שנה מעוברת.csv'
else:
    file_name = 'שנה רגילה.csv'


print("Hello and welcome to the Automated Synagogue Mail System Ver: 1.0.\n"
      "The mails will be sent today at " + TIME + ".\n"
      "If you don't wish to send the mails, please close this program.\n"
      "To change it, edit the 'settings.json' file.\n"
      "For more information, please open the 'Read Me' file.\n")

now = parser.parse(time.strftime("%H:%M:%S"))
send_time = parser.parse(TIME)
delta_time = send_time - now
while str(delta_time) != "0:00:00":
    time.sleep(1)
    now = parser.parse(time.strftime("%H:%M:%S"))
    send_time = parser.parse(TIME)
    delta_time = send_time - now

print("The program will now send the emails: \n")

with open(file_name) as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    next(csv_reader)
    for row in csv_reader:
        all_emails.append(row[6])
        if check_if_within_14_days_from_today(from_csv_to_date(heb_month[row[0]], heb_day[row[1]])):
            deceased_month.append(row[0])
            deceased_day.append(row[1])
            deceased_donor.append(row[2])
            deceased_relation.append(row[3])
            deceased_name.append(row[4])
            deceased_SoD.append(row[5])
            deceased_email.append(row[6])

all_emails = remove_duplicates(all_emails)
to_the_gizbar = all_emails
all_emails = list(filter(None, all_emails))

try:
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.ehlo()
    smtpObj.login(SENDER_EMAIL, SENDER_PASSWORD)
except Exception:
    print("Error: The email address or the password are incorrect. Please check them and try again.")
    input()
    sys.exit(0)

email_msg2 = ""

i = -1
for deceased in deceased_name:
    i = i + 1
    print("Yahrzeit for: {} \n".format(deceased_name[i]))
    subject = "יום הזיכרון של {}".format(deceased_name[i])
    with open(file_name) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        next(csv_reader)
        for email in all_emails:
            for row in csv_reader:
                month, day, donor, relation, deceased, SoD, email_flag = row
                msg = MASSAGE.format(donor, str(deceased_day[i] + " " + deceased_month[i]),
                      datetime.datetime.strptime(str(dates.HebrewDate.to_greg(from_csv_to_date(heb_month[deceased_month[i]], heb_day[deceased_day[i]]))), '%Y-%m-%d').strftime('%d/%m/%Y')
                      ,deceased_name[i], parshios.getparsha_string(dates.HebrewDate(year, heb_month[deceased_month[i]],
                      heb_day[deceased_day[i]])), deceased_name[i], deceased_SoD[i])
                email_msg = "נושא: {}\n\n {}".format(subject, msg)
                email_msg = MIMEText(email_msg, "plain", "utf-8")
                email_msg2 = email_msg
            time.sleep(3)
            print("Sending to: " + email)
            try:
                smtpObj.sendmail(SENDER_EMAIL, email, email_msg2.as_string().encode('ascii'))
            except Exception:
                print("Error: Email wasn't sent to {} - Please check if is written correctly and try again.".format(email))
                continue

update_the_gizbar = []
print("\n")
print("Updating the gizbar for those without emails...\n")

with open(file_name) as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    next(csv_reader)
    for row in csv_reader:
        month, day, donor, relation, deceased, SoD, email_flag = row
        if email_flag is "":
            update_the_gizbar.append(donor)

    update_the_gizbar = remove_duplicates(update_the_gizbar)
    update_the_gizbar = '\n'.join(list(update_the_gizbar))
    subject = "הנידון: נא לעדכן את האנשים חסרי המייל"
    email_msg2 = "שלום רב. לאנשים הבאים אין כתובת אימייל, אבקש לעדכן אותם אישית בנוגע ליארצייט החלים בשבועיים הקרובים:\n{}".format(update_the_gizbar)
    email_msg2 = MIMEText(email_msg2, "plain", "utf-8")
    email_msg2.as_string().encode('ascii')
    email_msg = "נושא: {} \n\n{}".format(subject, email_msg2)
try:
    smtpObj.sendmail(SENDER_EMAIL, SENDER_EMAIL, email_msg2.as_string().encode('ascii'))
except Exception:
    print("Error: Email wasn't sent to {} - Please check if is written correctly and try again.".format(SENDER_EMAIL))
smtpObj.quit()

print("All emails are successfully sent.")
