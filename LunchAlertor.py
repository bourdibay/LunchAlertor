
import win32com.client
import win32com
import os
import sys
import datetime
import ctypes  # An included library with Python install.
import argparse
import tkinter  as tk

parser = argparse.ArgumentParser()
parser.add_argument("startHour", type=str,
                    help="the start hour of the range we want to monitor. Example: 12")
parser.add_argument("endHour", type=str,
                    help="the end hour of the range we want to monitor. Example: 14")
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
startStr = start.strftime("%Y-%d-%m {0}:00".format(startHour))
endStr = start.strftime("%Y-%d-%m {0}:00".format(endHour))
appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"

restriction = "[Start] >= '{0}' AND [Start] <= '{1}'".format(startStr, endStr)
restrictedItems = appointments.Restrict(restriction)

# iterate over all meetings
listDisplayMsg = []
for appointmentItem in restrictedItems:
   startDateFormatted = appointmentItem.start.strftime("%H:%M")
   listDisplayMsg.append("[{0}] => {1}".format(startDateFormatted, appointmentItem.Subject))

class Window:
   def __init__(self, title, text):
      self.m_window = tk.Tk()
      self.m_window.configure(background="red")
      self.m_window.attributes('-fullscreen', True)
      self.m_label = tk.Label(self.m_window, text=text, font=("Courier bold", 44), width=400, anchor=tk.CENTER, bg='red', fg='white' )
      self.m_label.pack(fill='x', expand=True)

   def change_color(self):
       current_color_window = self.m_window.cget("background")
       current_color_fg_label = self.m_label.cget("foreground")
       next_color_window = "white" if current_color_window == "red" else "red"
       next_color_fg_label = "red" if current_color_fg_label == "white" else "white"
       self.m_window.config(background=next_color_window)
       self.m_label.config(background=next_color_window)
       self.m_label.config(foreground=next_color_fg_label)
       self.m_window.after(1000, self.change_color)

   def run(self):
      self.m_window.mainloop()

# display the meetings that we have to remind before going to lunch
if listDisplayMsg:
   title = "You have appointments between {0} and {1}".format(startHour, endHour)
   text = '\n'.join(listDisplayMsg)
   text = "Attention tu as des meetings bientôt\n\n" + text
   window = Window(title, text)
   window.change_color()
   window.run()
