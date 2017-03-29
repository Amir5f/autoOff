# Author: Amir Fischer
# Date: 29/3/2017
# App to shut down computer when idle for more than a configurable time.

# v2.0
# First call for idle time is suspended to avoid going into shutdown because long idle time (sleep)
# Prevent shutdown is based on moving the mouse during the alert time, not killing the task
# Added gui shutdown in case command didn't work

# v1.2
# Note: config.txt:
# enter a value in the field 'Max Idle Time:' to control the idle time
# enter True or False in the field 'Verbose:' to control the logging level
# Note: prevent shutdown process by killing (in task manager) app called autoOff.exe