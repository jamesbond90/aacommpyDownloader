import win32com.client
import os

# create a new task scheduler object
scheduler = win32com.client.Dispatch('Schedule.Service')

# connect to the local machine task scheduler
scheduler.Connect()

# create a new task scheduler task object
task = scheduler.NewTask(0)

# configure the task scheduler settings
task.RegistrationInfo.Description = 'Update AAComm'
task.Settings.Enabled = True
task.Settings.Hidden = False
task.Settings.StartWhenAvailable = True

# create a trigger for the task scheduler task
trigger = task.Triggers.Create(2) # 2 means monthly trigger
trigger.StartBoundary = '2023-03-05T20:00:00' # set the starting time
trigger.Id = 'MyTrigger'
trigger.Enabled = True
monthly = trigger.MonthlyTrigger # get the monthly trigger options
monthly.DaysOfMonth = [1] # run on the 1st day of the month

# create an action for the task scheduler task
cwd = os.getcwd()
action = task.Actions.Create(0)
action.Path = cwd + '\\Update.py'
action.Arguments = ''

# save the task scheduler task
root_folder = scheduler.GetFolder('\\')
root_folder.RegisterTaskDefinition(
    'MyTask',  # name of the task scheduler task
    task,  # task scheduler task object
    6,  # create or update existing task
    '',  # user account to run the task
    '',  # password for user account
    0)  # do not specify for interactive user account
