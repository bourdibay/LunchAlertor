
import win32com.client
import win32com
import os
import sys
import datetime
import ctypes  # An included library with Python install.
import argparse

def popupWarning(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

parser = argparse.ArgumentParser()
parser.add_argument("startHour", type=str,
                    help="the start hour of the range we want to monitor. Example: 12:00")
parser.add_argument("endHour", type=str,
                    help="the end hour of the range we want to monitor. Example: 14:00")
args = parser.parse_args()
startHour = args.startHour
endHour = args.endHour

try:
   app = win32com.client.GetActiveObject('Outlook.Application')
except:
   app = win32com.client.Dispatch('Outlook.Application')

outlook = app.GetNamespace("MAPI")
calendar = outlook.GetDefaultFolder(9)

appointments = calendar.Items

start = datetime.date.today()
# format can change in other regions
startStr = start.strftime("%Y-%d-%m {0}".format(startHour))
endStr = start.strftime("%Y-%d-%m {0}".format(endHour))
appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"

restriction = "[Start] >= '{0}' AND [Start] <= '{1}'".format(startStr, endStr)
restrictedItems = appointments.Restrict(restriction)

# iterate over all meetings
listDisplayMsg = []
for appointmentItem in restrictedItems:
   listDisplayMsg.append("At {0} => {1}".format(appointmentItem.start, appointmentItem.Subject))

# display the meetings that we have to remind before going to lunch
if listDisplayMsg:
   title = "You have appointments between {0} and {1}".format(startHour, endHour)
   text = '\n'.join(listDisplayMsg)
   popupWarning(title, text, 1)
