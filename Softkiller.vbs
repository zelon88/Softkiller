'File Name: Softkiller.vbs
'Version: v1.0, 10/28/2019
'Author: Justin Grimes, 10/28/2019

'Supported Arguments
  ' -e  (Email)  =  Set 'emailResult' config entry to TRUE (send emails when run).
  ' -o  (Output)  =  Set 'outputResult' config entry to TRUE (create a log file when this application kills other applications).
  ' -v  (verbose)  =  Set 'verbose' config entry to TRUE (log output to the console).
  ' -f  (Forced)  =  Set 'force' config entry to TRUE (bypass Office Application detection).
  ' -k  (Process To Kill)  =  Set '-l <process name>' to the complete name of a process to kill (required).
  ' -h  (Help)  =  Use the 'help' argument to display instructional text about this application.

' --------------------------------------------------
'Declare all variables to be used during execution of this application.
'Undeclared variables will cause a critical error and halt application execution.
Option Explicit
Dim argms, emailResult, outputResult, verbose, force, killExe, strComputer, strProgramToKill, SKScriptName, SKAppPath, SKLogPath, companyName, companyAbbr, companyDomain, _
 toEmail, SKMailFile, objFSO, objWMIService, strSafeDate, strSafeTime, strDateTime, logFileName, scriptPath, i, oFile, objlogFile, message, officeApp, officeApps, skip, _
 helpText, echoText, notText, methodText, killStatus, killResult, objShell, strUserName, objScript, strComputerName, colProcessList, objProcess, logFilePath, objApp
' --------------------------------------------------

' --------------------------------------------------
  ' ----------
  ' Company Specific variables.
  ' Change the following variables to match the details of your organization.
  
  ' The " SKScriptName" is the filename of this script.
  SKScriptName = "Softkiller.vbs"
  ' The "SKAppPath" is the full absolute path for the script directory, with trailing slash.
  SKAppPath = "\\SERVER\AutomationScripts\Softkiller\"
  ' The "SKLogPath" is the full absolute path for where network-wide logs are stored.
  SKLogPath = "\\SERVER\Logs\"
  ' The "companyName" the the full, unabbreviated name of your organization.
  companyName = "Company Inc."
  ' The "companyAbbr" is the abbreviated name of your organization.
  companyAbbr = "Company"
  ' The "companyDomain" is the domain to use for sending emails. Generated report emails will appear
  ' to have been sent by "COMPUTERNAME@domain.com"
  companyDomain = "company.com"
  ' The "toEmail" is a valid email address where notifications will be sent.
  toEmail = "IT@company.com"
  ' Set "emailResult" to TRUE to receive an email when registry modifications are detected. 
  ' Default is TRUE.
  emailResult = TRUE
  ' Set "outputResult" to TRUE to create a lot file when registry modifications are detected. 
  ' Default is TRUE.
  outputResult = TRUE
  ' When "outputResult" is set to TRUE, set "verbose" to TRUE to create a logfile on success or on error (default is error only).
  ' Default is FALSE.
  verbose = FALSE
  ' Set "force" to TRUE to force the script to continue even when it does not have elevated priviledges.
  ' Default is FALSE.
  force = FALSE
  ' ----------
' --------------------------------------------------

' --------------------------------------------------
'Set commonly used objects.
strComputer = "."
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objShell = CreateObject("Wscript.Shell")
Set argms = WScript.Arguments.Unnamed
'Some basic global variables.
officeApps = Array("WINWORD.EXE", "OUTLOOK.EXE", "EXCEL.EXE", "POWERPOINT.EXE")
strProgramToKill = ""
methodText = ""
echoText = ""
nottext = ""
killExe = ""
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
'Date/Time related variables.
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime
'File/Directory path related variables.
scriptPath = objFSO.GetParentFolderName(objScript) 
logFileName = SKLogPath & strComputerName & "-" & strDateTime & "-Softkiller.txt"
SKMailFile = "C:\Users\" & strUserName & "\Softkiller_Warning.mail"
' --------------------------------------------------

' --------------------------------------------------
'Retrieve the specified arguments.
  ' -e  (Email)  =  Set 'emailResult' config entry to TRUE (send emails when run).
  ' -o  (Output)  =  Set 'outputResult' config entry to TRUE (create a log file when this application kills other applications).
  ' -v  (verbose)  =  Set 'verbose' config entry to TRUE (log output to the console).
  ' -f  (Forced)  =  Set 'force' config entry to TRUE (bypass Office Application detection).
  ' -k  (Process To Kill)  =  Set '-l <process name>' to the complete name of a process to kill (required).
  ' -h  (Help)  =  Use the 'help' argument to display instructional text about this application.
Function ParseArgs()
  ParseArgs = FALSE
  'Iterate through all supplied arguments.
  For i = 0 to argms.Count -1
    'Detect the -e argument.
    If argms.item(i) = "-e" Then
      emailResult = TRUE
    End If
    'Detect the -o argument.
    If argms.item(i) = "-o" Then
      outputResult = TRUE
    End If
    'Detect the -v argument.
    If argms.item(i) = "-v" Then
      verbose = TRUE
    End If
    'Detect the -f argument.
    If argms.item(i) = "-f" Then
      force = TRUE
    End If
    'Detect the -h argument.
    'Displays help text.
    If argms.item(i) = "-h" Then
      helpText = "Usage:  " & SKScriptName & " -k <App-To-Kill.exe> -f -o -e -v" & VBNewLine & _
       " -e  (Email)  =  Set 'emailResult' config entry to TRUE (send emails when run)." & VBNewLine & _
       " -o  (Output)  =  Set 'outputResult' config entry to TRUE (create a log file when this application kills other applications)." & VBNewLine & _
       " -v  (verbose)  =  Set 'verbose' config entry to TRUE (log output to the console)." & VBNewLine & _
       " -f  (Forced)  =  Set 'force' config entry to TRUE (bypass Office Application detection)." & VBNewLine & _
       " -k  (Process To Kill)  =  Set '-l <process name>' to the complete name of a process to kill (required)." & VBNewLine & _
       " -h  (Help)  =  Use the 'help' argument to display instructional text about this application."
      WScript.Echo(helpText)
    End If
    'Detect the -k argument.
    'This is the only argument that is required for script execution.
    'Without a -k argument specified this script will not run.
    If argms.item(i) = "-k" Then
      killExe = argms.item(i + 1)
      ParseArgs = TRUE
    End If
  Next 
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create all required directories before the script can be run & delete any partial files that may already exist.
Function CreateReqdDirs()
  CreateReqdDirs = FALSE
  'Ensure a SKLogPath exists. Errors at this point probably indicate an intermediary directory does not exist or is not writable.
  If Not objFSO.FolderExists(SKLogPath) Then
    objFSO.CreateFolder(SKLogPath)
  End If
  'Double check to be sure that required folders were created. 
  If objFSO.FolderExists(SKLogPath) Then
    CreateReqdDirs = TRUE
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create a Warning.mail file. Use to prepare an email before calling sendEmail().
Function CreateEmail()
  'Check for an existing mail file and delete one if one exists.
  If objFSO.FileExists(SKMailFile) Then
    objFSO.DeleteFile(SKMailFile)
  End If
  'Check for an existing mail file and create one if none exists.
  If Not objFSO.FileExists(SKMailFile) Then
    objFSO.CreateTextFile(SKMailFile)
  End If
  'Set a handle for the "SKMailFile".
  Set oFile = objFSO.CreateTextFile(SKMailFile, True)
  'Write the actual email data to the mail file.
  oFile.Write "To: " & toEmail & vbNewLine & "From: " & strComputerName & "@" & companyDomain & vbNewLine & _
   "Subject: " & companyAbbr & " Softkiller Warning!!!" & vbNewLine & _
   "This is an automatic email from the " & companyName & " Network to notify you that an application was automatically killed." & _
   vbNewLine & vbNewLine & "Please verify that the equipment listed below is functioning properly." & vbNewLine & _
   vbNewLine & "USER NAME: " & strUserName & vbNewLine & "WORKSTATION: " & strComputerName & vbNewLine & "PROCESS TERMINATED: " & killExe & VBNewLine & "OPERATION RESULT: " & UCase(killStatus) & _
   vbNewLine & vbNewLine & "This check was generated by " & strComputerName & "." & vbNewLine & vbNewLine & "Script: """ & SKScriptName & """" 
   'Close the mail file.
  oFile.close
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function for running SendMail to send a prepared Warning.mail email message.
Function SendEmail() 
  objShell.run "c:\Windows\System32\cmd.exe /c sendmail.exe " & SKmailFile, 0, TRUE
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create a log file when -l is set.
'Returns "True" if logFilePath exists, "False" on error.
Function CreateSoftKillLog(message)
'Make sure the message is not blank.
  If message <> "" Then
    'Set a handle for the "logFileName".
    Set objlogFile = objFSO.CreateTextFile(logFileName, True)
    'Write the "message" to the log file.
    objlogFile.WriteLine(message)
    'Close the log file.
    objlogFile.Close
  End If
  'Check that a lot file was created and return the result.
  If objFSO.FileExists(logFilePath) Then
    error = FALSE
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to detect if the selected application is running.
'If the selected application to kill is a Microsoft Office application then we can access it with CreateObject.
'Killing a Microsoft Office application using "objApp.Quit" is much gentler than using "objApp.Terminate()."
'If we simply terminate a program like Outlook while the PST's are being accessed we might corrupt data.
'Returns "TRUE" on success. Returns "FALSE" on error.
Function KillProcess(strProgramToKill)
    KillProcess = FALSE
    skip = FALSE
    'Execute the query set in global variables and return processes which match the user supplied application.
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProgramToKill & "'")
    'Iterate through the results of the "colProcessList" query and return any process matching the one supplied by the user.
    For Each objProcess in colProcessList
      'Loop through each element in the "officeApps" array for each process found in the loop above.
      For Each officeApp in officeApps
        'See if the current process is a match for the one specified by the user.
        If LCase(strProgramToKill) = LCase(objProcess.Name) Then
          'See if the current process is in the "officeApps" array.
          If LCase(objProcess.Name) = LCase(officeApp) Then
            Set objApp = CreateObject(Replace(officeApp, ".EXE", "") & ".Application") 
            'Kill the selected Office application.
            objApp.Quit
            KillProcess = TRUE
            skip = TRUE
            methodText = " gently"
            WScript.Sleep(5)
          End If
        End If
      Next
      'If the "force" argument is set we terminate the currently selected program at the end of the script regardless.
      If Not skip And force Then
        objProcess.Terminate()
        KillProcess = TRUE
      End If
      If Not skip And Not force Then
        'Termination the currently selected process.
        objProcess.Terminate()
        KillProcess = TRUE
      End If
    Next
    'Execute the query set again to check that the terminated process was actually terminated.
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProgramToKill & "'")
    'Iterate through the results of the "colProcessList" query and return any process matching the one supplied by the user.
    For Each objProcess in colProcessList
      'See if the current process is still in the "officeApps" array.
      If LCase(objProcess.Name) = LCase(officeApp) Then
        KillProcess = FALSE
      End If
    Next
    'Prepare some text to use for console & log entries.
    If KillProcess = FALSE Then
      killStatus = "Failed"
      notText = "not "
    Else 
      killStatus = "Succeeded"
    End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'The main logic & entry point for the script. Makes use of the functions above.

'Parse the arguments supplied to the script and use them to prepare the operating environment for the session.
'If no arguments are supplied hard-coded configuration entries will be used instead.
If ParseArgs() Then
  'Create directories & verify user input is valid.
  If CreateReqdDirs() and Len("" & killExe) > 0 Then
    killResult = KillProcess(killExe)
    echoText = Replace(SKScriptName, ".vbs", "") & ", " & strDateTime & ": Operation " & killStatus & "! " & killExe & " was " & notText & "terminated" & methodText & "." 
  End If
  'Send an email if the "-e" argument or config entry is set.
  If emailResult Then
    CreateEmail()
    SendEmail()
  End If
  'Create a log file if the "-o" argument or config entry is set.
  If outputResult Then
    CreateSoftKillLog(echoText)
  End If
  'Write output to the console if the "-v" argument or congig entry is set.
  If verbose Then
    WScript.Echo(echoText)
  End If
End If
' --------------------------------------------------
