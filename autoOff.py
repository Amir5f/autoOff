# Author: Amir Fischer
# Date: 29/3/2017
# App to shut down computer when idle for more than 90 minutes
# v2.0
# First call for idle time is suspended to avoid going into shutdown because long idle time (sleep)
# Prevent shutdown is based on moving the mouse during the alert time, not killing the task
# Added gui shutdown in case command didn't work
# v1.2
# Note: config.txt:
# enter a value in the field 'Max Idle Time:' to control the idle time
# enter True or False in the field 'Verbose:' to control the logging level
# Note: prevent shutdown process by killing (in task manager) app called autoOff.exe

import os
import pywinauto
import time
import win32api
import win32con
from ctypes import Structure, windll, c_uint, sizeof, byref

MAX_TIME = 90        # 90 minutes
INTERVALS = 10*60            # 10 minutes
WAIT = 5*60                  # 5 minutes

DEBUG_MODE = False


# Return the number of minutes since the last time the computer received any input (mouse/keyboard)
def time_since_last_use():
    lastInputInfo = LASTINPUTINFO()
    lastInputInfo.cbSize = sizeof(lastInputInfo)
    windll.user32.GetLastInputInfo(byref(lastInputInfo))
    millis = windll.kernel32.GetTickCount() - lastInputInfo.dwTime
    minutes = millis / 1000.00 / 60.00
    return minutes


class LASTINPUTINFO(Structure):
    _fields_ = [
        ('cbSize', c_uint),
        ('dwTime', c_uint)
    ]


# Opens notepad and prints a message about shutdown and then exists it without saving.
def alert_shutdown():
    delay = 100

    # Inform
    app = pywinauto.Application().start(r"c:\Windows\System32\notepad.exe")

    app.Notepad.Edit.TypeKeys("Computer was not touched for ", with_spaces=True)
    app.Notepad.Edit.TypeKeys(MAX_TIME, with_spaces=True)
    app.Notepad.Edit.TypeKeys(" minutes and will shutdown in ", with_spaces=True)
    app.Notepad.Edit.TypeKeys(WAIT, with_spaces=True)
    app.Notepad.Edit.TypeKeys("minutes.", with_spaces=True)

    # Note
    app.Notepad.TypeKeys("{ENTER}")
    app.Notepad.Edit.TypeKeys("Note: to prevent shutdown move the mouse in the next ", with_spaces=True)
    app.Notepad.Edit.TypeKeys(WAIT, with_spaces=True)
    app.Notepad.Edit.TypeKeys("minutes.", with_spaces=True)

    # Delay/chance to prevent shutdown
    time.sleep(WAIT)
    # If the shutdown was prevented
    if time_since_last_use() < MAX_TIME:
        app.Notepad.TypeKeys("{ENTER}")
        app.Notepad.Edit.TypeKeys("OK, your'e still here...")
        time.sleep(4)
        app.Notepad.MenuSelect("File -> Exit -> Don't Save")
        app.Notepad.TypeKeys("{TAB}{SPACE}")

        # go back to monitoring idleness
        if DEBUG_MODE:
            idleness_check_debug(MAX_TIME)
        else:
            idleness_check(MAX_TIME)

    # If the shutdown was not prevented
    else:
        app.Notepad.TypeKeys("{ENTER}{ENTER}")
        app.Notepad.Edit.TypeKeys("Shutdown...")
        time.sleep(delay)

        # Exit Notepad
        app.Notepad.MenuSelect("File -> Exit -> Don't Save")
        app.Notepad.TypeKeys("{TAB}{SPACE}")


# Logs the current time and sends the computer to shut down
def initiate_shutdown():
    timeStamp = get_time()
    msg = timeStamp + "\t shutdown"
    log(msg)

    os.system("shutdown -s ")
    # In case shutdown was not successful, use GUI:
    gui_shutdown()


def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)


def gui_shutdown():
    click(20, 1030)
    time.sleep(1)
    click(20, 980)
    time.sleep(1)
    # click(20, 930)


# Reads the configuration file and determines the allowed idle time
# If the config file is missing - default value is set
def config_time():

    try:
        config_file = open("config.txt", "r")
    except getattr(__builtins__,'FileNotFoundError', IOError):
        return

    line = config_file.readline()
    while line:
        if 'Max Idle Time: ' in line:
            val = line.split(': ', 1)
            MAX_TIME = int(val[1])
            return

        line = config_file.readline()

    return


# Reads the configuration file and determines the logging level that is specified.
# If the config file is missing - regular debugging is done
def config_debug():

    try:
        config_file = open("config.txt", "r")
    except getattr(__builtins__,'FileNotFoundError', IOError):
        return

    line = config_file.readline()
    while line:
        if 'Verbose: ' in line:
            if 'True' in line:
                DEBUG_MODE = True

        line = config_file.readline()
    return


# Prints msg to the log file
def log(msg):

    log_file = open("log.txt", "a")
    log_file.write(msg)
    log_file.close()


# Returns string with the local time, for logs
def get_time():
    localtime = time.asctime(time.localtime(time.time()))
    localtime = "\n"+localtime

    return localtime


# Prints to the log file that an idleness check was done (verbose mode)
def log_idle_check(time_idle):
    timeStamp = get_time()
    msg = timeStamp + "\t idle for " + time_idle
    log(msg)


# Performs the idleness check and prints res only to terminal
def idleness_check():
    idle_time = 0
    print("Debug mode: False")

    while idle_time < MAX_TIME:
        idle_time = round(time_since_last_use(), 2)
        print("idle for " + str(idle_time) + " minutes")

        time.sleep(INTERVALS)
        continue


# Performs the idleness check and prints res to terminal and to log (verbose)
def idleness_check_debug():
    idle_time = 0
    print("Debug mode: True")

    while idle_time < MAX_TIME:
        idle_time = round(time_since_last_use(), 2)
        print("idle for " + str(idle_time) + " minutes")
        log_idle_check(str(idle_time))

        time.sleep(INTERVALS)
        continue

    alert_shutdown()

    initiate_shutdown()


def main():

    config_time()
    print("Max idle time configured: ", MAX_TIME)

    # Delay run to avoid going into shutdown process because the computer responds as idle on the time it was off
    time.sleep(INTERVALS)

    if DEBUG_MODE:
        idleness_check_debug()
    else:
        idleness_check()


if __name__ == "__main__":
    main()
