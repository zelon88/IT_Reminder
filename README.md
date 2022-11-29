NAME: IT_Reminder.vbs

TYPE: Visual Basic Script

PRIMARY LANGUAGE: VBS

AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 1/8/2019

CURRENT VERSION DATE: 11/28/2022

VERSION: v2.2

DESCRIPTION: 
A simple script for reminding the IT department about due maintenance items.


PURPOSE: 
To ensure NIST/PCI compliance as well as ensure that maintenance is performed regularly and properly. To create and maintain maintenance logs in an automated fashion.


INSTALLATION INSTRUCTIONS: 
1. Copy the entire "IT_Reminder" folder into the "AutomationScripts" folder on SERVER (or any other network accesbible location).
2. Add a scheduled task to SERVER to run "IT_Reminder.vbs" once per day with the "-t" argument.
3. Use the -t argument when running from Task Scheduler. This will send emails to IT warning them about due maintenance.
4. Do not use the -t argument when running as a user. This will allow maintenance records to be updated.
5. Open IT_Reminder.vbs with a text editor and modify the global variables at the start of the script to match your environment.
6. Modify the checkArr array so that each first index matches a .html file from the "/Data" directory.
7. A correct configuration entry will look like   checkArr(0,0) = "Name_Of_Your_Maint_Procedure"
8. Once you have re-labelled the checkArr array to match the names of your maintanence procedures, locate the corresponding .html file from the "/Data" directory.
9. Add the maintanence procedure to the .html which corresponds to the checkArr maintanence item you are adding.
10. When each maintanence item is due, the .html file containing your maintanence procedure will be displayed and sent to the user.
11. When maintanence is recorded by a user; the recorded data is stored in the maintDir as an Excel spreadsheet with a name which corresponds to the maintanence item being performed. 

NOTES: 
1. SendMail for Windows is required and included in the "IT_Reminder" folder. 
2. The SendMail data files must be included in the same directory as "IT_Reminder.vbs" in order for emails to be sent correctly.

SUPPORTED SWITCHES
 -t  -Use with task scheduler to disable console output.
 -u  -Use to display upcoming maintenence without updating cache or records.
 -s  -Use to display due maintenance without updating cache or records.