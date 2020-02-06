'File Name: IT_Reminder.vbs
'Version: v2.1, 2/6/2020
'Author: Justin Grimes, 7/16/2018

'--------------------------------------------------
'Supported Switches
' -t  -Use with task scheduler to disable console output.
' -u  -Use to display upcoming maintenence without updating cache or records.
' -s  -Use to display due maintenance without updating cache or records.
'--------------------------------------------------

'--------------------------------------------------
'Global variable definitions used in this script.
Option Explicit
dim strComputer, param1, objFSO, objWMIService, arg, objNet, objShell, strComputerName, strUserName, timeStamp, objExcel, _
mailFile, logFile, cacheFile, checkArr(19,2), cacheDir, checkFile, checkFrequency, today, maxAge, strNotes, _
dFileMod, triggerEmail, writeFile, checkDataFile, dataFile, data, mFile, dataDir, maintDir, check, counter, _
checkFile1, checkFile2, appexcel, wb, maintFile, objWorkbook, r, stat, dateDue, companyAbbr, toEmail, fromEmail
'--------------------------------------------------

'--------------------------------------------------
'Define variables for the session.
counter = 0
strComputer = "."
param1 = checkDataFile = strNotes = dataFile = ""
cacheDir = "\\Server\Scripts\IT_Reminder\Cache"
dataDir = "\\Server\Scripts\IT_Reminder\Data"
maintDir = "\\Server\IT\IT Resources\Maintenance"
companyAbbr = "Company"
toEmail = "IT@Company.com"
fromEmail = "Server@Company.com"
today = Date
timeStamp = Now
maxAge = 0
triggerEmail = FALSE
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set arg = WScript.Arguments
Set objNet = CreateObject("Wscript.Network") 
Set objShell = WScript.CreateObject("WScript.Shell")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
mailFile = "Warning.mail"
'--------------------------------------------------

'--------------------------------------------------
'Define array elements for the session.
'This check is for keeping visibility on the logs directory.
'We build an array of checks to perform that will later be used to include .html files 
  'of the same names. The first array entry is the name of the check to be performed, and the 
  'secont entry is the frequency for the check. Supported frequencies are "DAILY", "WEEKLY", 
  '"MONTHLY", and "YEARLY"
'If you want to add or remove items from this array, you must also remember to modify the array indecies defined 
  'in the global variables section of this script.
'This check is for . . .
checkArr(0,0) = "DAILY_MAINT_ITEM_1"
checkArr(0,1) = "DAILY"
'This check is for . . .
checkArr(1,0) = "DAILY_MAINT_ITEM_2"
checkArr(1,1) = "DAILY"
'This check is for . . .
checkArr(2,0) = "DAILY_MAINT_ITEM_3"
checkArr(2,1) = "DAILY"
'This check is for . . .
checkArr(3,0) = "DAILY_MAINT_ITEM_4"
checkArr(3,1) = "DAILY"
'This check is for . . .
checkArr(4,0) = "WEEKLY_MAINT_ITEM_1"
checkArr(4,1) = "WEEKLY"
'This check is for . . .
checkArr(5,0) = "WEEKLY_MAINT_ITEM_2"
checkArr(5,1) = "WEEKLY"
'This check is for . . .
checkArr(6,0) = "WEEKLY_MAINT_ITEM_3"
checkArr(6,1) = "WEEKLY"
'This check is for . . .
checkArr(7,0) = "WEEKLY_MAINT_ITEM_4"
checkArr(7,1) = "WEEKLY"
'This check is for . . .
checkArr(8,0) = "MONTHLY_MAINT_ITEM_1"
checkArr(8,1) = "MONTHLY"
'This check is for . . .
checkArr(9,0) = "MONTHLY_MAINT_ITEM_2"
checkArr(9,1) = "MONTHLY"
'This check is for . . .
checkArr(10,0) = "MONTHLY_MAINT_ITEM_3"
checkArr(10,1) = "MONTHLY"
'This check is for . . .
checkArr(11,0) = "MONTHLY_MAINT_ITEM_4"
checkArr(11,1) = "MONTHLY"
'This check is for . . .
checkArr(12,0) = "MONTHLY_MAINT_ITEM_5"
checkArr(12,1) = "MONTHLY"
'This check is for . . .
checkArr(13,0) = "MONTHLY_MAINT_ITEM_6"
checkArr(13,1) = "MONTHLY"
'This check is for . . .
checkArr(14,0) = "MONTHLY_MAINT_ITEM_7"
checkArr(14,1) = "MONTHLY"
'This check is for . . .
checkArr(15,0) = "MONTHLY_MAINT_ITEM_8"
checkArr(15,1) = "MONTHLY"
'This check is for . . .
checkArr(16,0) = "MONTHLY_MAINT_ITEM_9"
checkArr(16,1) = "MONTHLY"
'This check is for . . .
checkArr(17,0) = "YEARLY_MAINT_ITEM_1"
checkArr(17,1) = "YEARLY"
'This check is for . . .
checkArr(18,0) = "YEARLY_MAINT_ITEM_2"
checkArr(18,1) = "YEARLY"
'--------------------------------------------------

'--------------------------------------------------
'A funciton for running SendMail.
Function SendEmail() 
  objShell.run "sendmail.exe " & mailFile
End Function
'--------------------------------------------------

'--------------------------------------------------
'The following code checks for required directories and creates them if needed.
If Not objFSO.FolderExists(maintDir) Then
  Set objFolder = objFSO.CreateFolder(maintDir)
End If
If Not objFSO.FolderExists(cacheDir) Then
  Set objFolder = objFSO.CreateFolder(cachedir)
End If
If Not objFSO.FolderExists(dataDir) Then
  Set objFolder = objFSO.CreateFolder(dataDir)
End If
'--------------------------------------------------

'--------------------------------------------------
'Retrieve the specified arguments.
If (arg.Count > 0) Then
  param1 = arg(0)
End If
'--------------------------------------------------

'--------------------------------------------------
'The following code is performed when the -t argument is set (intended for scheduled tasks).
  'When run as a task there are no outputs to the command line or dialouge boxes. Instead
  'an email is sent to IT@company.com for every maintenance item that is due. 
If (param1 = "-t") Then
  For Each check In checkArr
    If counter > 19 Then
      Exit For
    End If
    If check = "" Then
      Exit For
    End If
    triggerEmail = False
    'The checkFile contains the maintenance log data.
    'The checkDataFile contains the html formatted instructions for performing the maintenance.
    checkFile = cacheDir & "\" & checkArr(counter,0) & ".txt"
    checkDataFile = dataDir & "\" & checkArr(counter,0) & ".html"
    checkFrequency = checkArr(counter,1)
    'Create the cache file if it was expired, deleted, or missing.
    If Not objFSO.FileExists(checkFile) Then
      maxAge = -1
    End If
    'If the checkFile exists, get it's age.
    If objFSO.FileExists(checkFile) Then
      Set checkFile1 = objFSO.GetFile(checkFile)
      dFileMod = FormatDateTime(checkFile1.DateLastModified, "2")
      'Set the maxAge if the current check is to be performed daily.
      If (checkFrequency = "DAILY") Then
        maxAge = 1
      End If
      'Set the maxAge if the current check is to be performed weekly.
      If (checkFrequency = "WEEKLY") Then
        maxAge = 7
      End If
       'Set the maxAge if the current check is to be performed monthly.
      If (checkFrequency = "MONTHLY") Then
        maxAge = 30
      End If
       'Set the maxAge if the current check is to be performed yearly.
      If (checkFrequency = "YEARLY") Then
        maxAge = 365
      End If
    End If
    'See if the cache file has expired.
    If DateDiff("d", dFileMod, today) >= maxAge Then
      triggerEmail = TRUE
    End If
    'Send an email if the selected cache file was expired.
    If (triggerEmail = TRUE) Then
      Wscript.Sleep(1000)
      Set dataFile = objFSO.OpenTextFile(checkDataFile, 1)
      data = dataFile.ReadAll
      dataFile.Close
      Set mFile = objFSO.CreateTextFile(mailFile, TRUE, FALSE)  
      mFile.Write "To: " & toEmail & vbNewLine & "From: " & fromEmail & vbNewLine & "Subject: " & companyAbbr & " IT Reminder, " & _
       checkArr(counter,0) & "!!!" & vbNewLine & data & vbNewLine & vbNewLine & _
       "This reminder was generated by " & strComputerName & " and is run once per day." & _
       vbNewLine & vbNewLine & "Script: ""IT_Reminder.vbs""" 
      mFile.Close
      Wscript.Sleep(1000)
      sendEmail()
      triggerEmail = FALSE
    End If
    counter = counter + 1
  Next
End If
'--------------------------------------------------

'--------------------------------------------------
'The following code is performed when the -t argument is NOT set (intended for users to run).
  'When run directly dialogue boxes will be displayed to prompt the user to input maintenance records.
If (param1 <> "-t") And (param1 <> "-u") And (param1 <> "-s") Then
  MsgBox "This application will check for due network maintenance. When maintenance is due this application will provide " & _ 
   "procedures and instructions for performing it. " & VBNewLine & VBNewLine & "This application also gathers information about " & _ 
   "performed maintainance and stores it in the form of Maintenance Record spreadsheets to ensure NIST compliance.", 64, companyAbbr & " IT Reminders"
  For Each check In checkArr
    If counter > 18 Then
      Exit For
    End If
    'The checkFile contains the maintenance log data.
    'The checkDataFile contains the html formatted instructions for performing the maintenance.
    checkFile = cacheDir & "\" & checkArr(counter,0) & ".txt"
    checkDataFile = dataDir & "\" & checkArr(counter,0) & ".html"
    maintFile = maintDir & "\" & checkArr(counter,0) & ".xlsx"
    checkFrequency = checkArr(counter,1)
    'If the checkFile exists, get it's age.
    If objFSO.FileExists(checkFile) Then
      Set checkFile1 = objFSO.GetFile(checkFile)
      dFileMod = FormatDateTime(checkFile1.DateLastModified, "2")
      'Set the maxAge if the current check is to be performed daily.
      If (checkFrequency = "DAILY") Then
        maxAge = 1
      End If
      'Set the maxAge if the current check is to be performed weekly.
      If (checkFrequency = "WEEKLY") Then
        maxAge = 7
      End If
       'Set the maxAge if the current check is to be performed monthly.
      If (checkFrequency = "MONTHLY") Then
        maxAge = 30
      End If
       'Set the maxAge if the current check is to be performed yearly.
      If (checkFrequency = "YEARLY") Then
        maxAge = 365
      End If
    End If
    If Not objFSO.FileExists(checkFile) Then
      maxAge = -1
    End If
    'See if the cache file has expired.
    If DateDiff("d", dFileMod, today) >= maxAge Then
      Set dataFile = objFSO.OpenTextFile(checkDataFile, 1)
      data = dataFile.ReadAll
      dataFile.Close
      'Throw a dialogue box to gather user input for the notes column of the maintenance record spreadsheet.
      MsgBox data, 64, checkArr(counter,0)
      strNotes = InputBox(checkArr(counter,0) & VBNewLine & vbNewLine & "Please perform the specified " & _
       "maintenance and enter your notes below." & VBNewLine & vbNewLine & "You MUST leave a note to mark this " & _
       "maintanence complete." & vbNewLine, checkArr(counter,0))
      If strNotes <> "" Then 
        'Recreate the cache file if it already exists.
        If objFSO.FileExists(checkFile) Then
          objFSO.Deletefile(checkFile)
          Set writeFile = objFSO.CreateTextFile(checkFile)
          writeFile.Close
        End If
        'Create the cache file if it was expired, deleted, or missing.
        If Not objFSO.FileExists(checkFile) Then
          Set writeFile = objFSO.CreateTextFile(checkFile)
          writeFile.Close
        End If
        'Create the maintenance file using Excel if it was deleted, or missing.
        If Not objFSO.FileExists(maintFile) Then
          Set objExcel = CreateObject("Excel.Application")
          objExcel.Visible = False
          Set objWorkbook = objExcel.Workbooks.Add()
          objWorkbook.SaveAs(maintFile)
          objExcel.Quit
        End If
        'If the maintenance file exists, append the users input to it.
        If objFSO.FileExists(maintFile) Then
          Set checkFile2 = objFSO.GetFile(maintFile)
          dFileMod = FormatDateTime(checkFile2.DateLastModified, "2")
          Set appexcel = WScript.CreateObject("Excel.Application")
          With appexcel
            .Visible = False
            Set wb = .Workbooks.Open(maintFile)
            r = 1
            Do Until Len(.Cells(r, 1).Value) = 0
              r = r + 1
            Loop
            .Cells(r, 1).Value = timeStamp
            .Cells(r, 2).Value = strUserName
            .Cells(r, 3).Value = strComputerName
            .Cells(r, 4).Value = checkArr(counter,0)
            .Cells(r, 5).Value = checkArr(counter,1)
            .Cells(r, 6).Value = strNotes
            wb.Save
            Set wb = Nothing
            .Quit
          End With
          Set appexcel = Nothing
        End If
      End If
    End If
    counter = counter + 1
  Next
  MsgBox "All maintenance complete!" & VBNewLine, 64, companyAbbr & " IT Reminders"
End If
'--------------------------------------------------

'--------------------------------------------------
' The following code is performed when the "-u" argument is set.
  'The "-U" argument displays upcoming maintenance but does not update the cache or records.
If (param1 = "-u") Then
  MsgBox "This application will check for upcoming network maintenance. When maintenance is due this application will provide " & _ 
   "procedures and instructions for performing it. " & VBNewLine & VBNewLine & "This application also gathers information about " & _ 
   "performed maintainance and stores it in the form of Maintenance Record spreadsheets to ensure NIST compliance.", 64, companyAbbr & " IT Reminders"
  For Each check In checkArr
    If counter > 18 Then
      Exit For
    End If
    'The checkFile contains the maintenance log data.
    checkFile = cacheDir & "\" & checkArr(counter,0) & ".txt"
    checkFrequency = checkArr(counter,1)
    'Create the cache file if it was expired, deleted, or missing.
    If Not objFSO.FileExists(checkFile) Then
      Set writeFile = objFSO.CreateTextFile(checkFile)
      writeFile.Close
      maxAge = -1
    End If
    'If the checkFile exists, get it's age.
    If objFSO.FileExists(checkFile) Then
      Set checkFile1 = objFSO.GetFile(checkFile)
      dFileMod = FormatDateTime(checkFile1.DateLastModified, "2")
      'Set the maxAge if the current check is to be performed daily.
      If (checkFrequency = "DAILY") Then
        maxAge = 1
      End If
      'Set the maxAge if the current check is to be performed weekly.
      If (checkFrequency = "WEEKLY") Then
        maxAge = 7
      End If
       'Set the maxAge if the current check is to be performed monthly.
      If (checkFrequency = "MONTHLY") Then
        maxAge = 30
      End If
       'Set the maxAge if the current check is to be performed yearly.
      If (checkFrequency = "YEARLY") Then
        maxAge = 365
      End If
    End If
    'See when the cache files expire and adjust the display message accordingly.
    If DateDiff("d", dFileMod, today) >= maxAge Then
      stat = "Overdue"
      If maxAge = -1 Then
        dFileMod = "Never"
        dateDue = "ASAP"
      End If
      If maxAge <> -1 Then
        dateDue = DateAdd("d", maxAge, dFileMod)
      End If
    End If
    If DateDiff("d", dFileMod, today) <= maxAge Then
      stat = "Current"
      dateDue = DateAdd("d", maxAge, dFileMod)
    End If
    MsgBox checkArr(counter,0) & " - " & checkArr(counter, 1) & VBNewLine & VBNewLine & "Interval: " & checkArr(counter, 1) & vbNewLine & _
     "Status: " & stat & vbNewLine & "Last Completion Date: " & dFileMod & VBNewLine & "Next Due Date: " & dateDue & vbNewLine & vbNewLine, 64, companyAbbr & " IT Reminders"
    counter = counter + 1
  Next
End If
'--------------------------------------------------

'--------------------------------------------------
'The following code is performed when the -t argument is NOT set (intended for users to run).
  'When run directly dialogue boxes will be displayed to prompt the user to input maintenance records.
If (param1 = "-s") Then
  MsgBox "This application will check for due network maintenance. When maintenance is due this application will provide " & _ 
   "procedures and instructions for performing it. ", 64, companyAbbr & " IT Reminders"
  For Each check In checkArr
    If counter > 18 Then
      Exit For
    End If
    'The checkFile contains the maintenance log data.
    'The checkDataFile contains the html formatted instructions for performing the maintenance.
    checkFile = cacheDir & "\" & checkArr(counter,0) & ".txt"
    checkDataFile = dataDir & "\" & checkArr(counter,0) & ".html"
    maintFile = maintDir & "\" & checkArr(counter,0) & ".xlsx"
    checkFrequency = checkArr(counter,1)
    'If the checkFile exists, get it's age.
    If objFSO.FileExists(checkFile) Then
      Set checkFile1 = objFSO.GetFile(checkFile)
      dFileMod = FormatDateTime(checkFile1.DateLastModified, "2")
      'Set the maxAge if the current check is to be performed daily.
      If (checkFrequency = "DAILY") Then
        maxAge = 1
      End If
      'Set the maxAge if the current check is to be performed weekly.
      If (checkFrequency = "WEEKLY") Then
        maxAge = 7
      End If
       'Set the maxAge if the current check is to be performed monthly.
      If (checkFrequency = "MONTHLY") Then
        maxAge = 30
      End If
       'Set the maxAge if the current check is to be performed yearly.
      If (checkFrequency = "YEARLY") Then
        maxAge = 365
      End If
    End If
    If Not objFSO.FileExists(checkFile) Then
      maxAge = -1
    End If
    'See if the cache file has expired.
    If DateDiff("d", dFileMod, today) >= maxAge Then
      Set dataFile = objFSO.OpenTextFile(checkDataFile, 1)
      data = dataFile.ReadAll
      'Throw a dialogue box to gather user input for the notes column of the maintenance record spreadsheet.
      MsgBox data, 64, checkArr(counter,0)
      dataFile.Close
    End If
    counter = counter + 1
  Next
  MsgBox "Maintenance scan complete!" & VBNewLine & VBNewLine & "No more due maintenance items to display!" & VBNewLine, 64, companyAbbr & " IT Reminders"
End If
'--------------------------------------------------