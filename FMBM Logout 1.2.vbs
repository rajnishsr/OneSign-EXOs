'FMBM Logout Version 1.2 date 12/22/2021
'Added functionality to check at the file size at share and Launchpad and if one of them is 0 KB, code quits

Sub Main

	DIM objFileToRead, FMBMFile, strLine, iCounter
	DIM userRoleFileAtShare, userRoleFileAtShareLM, userRoleFileLM, a, b, c, LogFile, LPRefreshCmd, TypeKey, FMBMLog, FMBMFileSize
	DIM fso, loggedinUser, shareLocation, roleQuery, userRole, userRoleFile, AgentType, FMBMLoggedInUser, FMBMLoggedInUserFile
	DIM ErrorBuffer
	
	SET fso=CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	
	'Error Buffer helps in collecting error messages till Error file is created.
	'Once FMBMLog is created, Error Buffer will write to that file
	ErrorBuffer =""


	FMBMLoggedInUser = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\LoggedInUser.txt"
	LPRefreshCmd = """C:\Program Files (x86)\Imprivata\OneSign Agent\x64\LaunchpadRefresh.cmd"""		

	'If FMBM folder do not exist, exit

	IF Not fso.FolderExists("C:\ProgramData\Launchpad\INIToolBackup\FMBM\") Then
		Wscript.Quit
	END IF	
		
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Logout - " & loggedinUser & ": Error while checking for FMBM folder: " & Err.Description & vbCrLf
		Err.Clear
	End If
	
	'Check for logged in user 3/4
	If fso.FileExists(FMBMLoggedInUser) Then
	'If file exists read shareLocation, roleQuery and userRoleFile from the file
		Set FMBMLoggedInUserFile = fso.OpenTextFile(FMBMLoggedInUser,1, false, -2)
		'LogFile.WriteLine(Now() & ":  " & loggedinUser &  ": FMBM property file exist " & FMBMFile)
		do while not FMBMLoggedInUserFile.AtEndOfStream
			strLine = FMBMLoggedInUserFile.ReadLine()
			strLine = Trim(strLine)
		 
			If Left(strLine, 1) =";" or Left(strLine, 1) = "#" Then
				'Yep, this string is commented!
			Else 
				loggedinUser = strLine
				ErrorBuffer = ErrorBuffer  & Now() & ":  Logout - " & loggedinUser & " <--Logged in User"  & vbCrLf
			End If
		loop
	End If
	
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Logout - " & loggedinUser & ": Error while checking for logged in User :" & Err.Description & vbCrLf
		Err.Clear
	End If
	
	FMBMFile = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\" &"FMBM_" & loggedinUser & ".txt"
	FMBMLog = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMLog.txt"	
	
	'Check if the Log file exist and is less then 5 MB
	If fso.FileExists(FMBMLog) Then
		Set FMBMFileSize = fso.GetFile(FMBMLog)
		If FMBMFileSize.Size > 5000000 Then
			FMBMFileSize.Delete(True)
		End If
		Set LogFile = fso.OpenTextFile(FMBMLog, 8, true)
	Else
		Set LogFile = fso.CreateTextFile(FMBMLog, true)
	End If
	
	'Write whatever is in ErrorBuffer
	LogFile.Write ErrorBuffer
	
	'Check if the FMBM file for the user exists. This file has all the config for the user.
	If fso.FileExists(FMBMFile) Then
	'If file exists read shareLocation, roleQuery and userRoleFile from the file
		Set objFileToRead = fso.OpenTextFile(FMBMFile,1)
		LogFile.WriteLine(Now()  & ":  Logout - " & loggedinUser &  ": FMBM property file exist " & FMBMFile)
		iCounter = 1
		do while not objFileToRead.AtEndOfStream
			strLine = objFileToRead.ReadLine()
			strLine = Trim(strLine)
		 
			If Left(strLine, 1) =";" or Left(strLine, 1) = "#" Then
				'Yep, this string is commented!
			Else 
				Select Case iCounter
					case 1 
						shareLocation = strLine
					case 2 
						roleQuery = strLine
					case 3 
						userRoleFile = strLine
					case 4
						AgentType = strLine
				End Select		
				iCounter= iCounter + 1
			End If
		loop
	Else 
		'If the FMBM property file do not exist, exit code.
		LogFile.WriteLine(Now() &  ":  Logout - " & loggedinUser & ": Exiting code as FMBM property file do not exist")
		Wscript.Quit
	End If
	
	userRoleFileAtShare = shareLocation & loggedinUser &"_" & userRoleFile

	'If User Role file exists at share, compare it with the Launchpad folder.
	'Whichever is newer is copied to the other location.
	'If the fie does not exist, copy the file from Launchpad folder to Role share folder.
	If fso.FileExists(userRoleFileAtShare) Then 
		Set b = fso.GetFile(userRoleFileAtShare)
		LogFile.WriteLine(Now() & ":  Logout - " & loggedinUser & ": User Role file at share exists and is " & userRoleFileAtShare)
		userRoleFileAtShareLM = b.DateLastModified
		Set c = fso.GetFile("C:\ProgramData\Launchpad\" & userRoleFile)
		userRoleFileLM = c.DateLastModified
		
		'Check to see if the role file at share or local is 0 KB, if it is then quit
		If(b.size = 0 or c.size = 0) Then
			LogFile.WriteLine(Now() & ":  Logout - " & loggedinUser &  ": Exiting code as file size is 0")
			LogFile.WriteLine(Now() & ":  Logout - " & loggedinUser &  ": Share " & userRoleFileAtShareLM & " size = " & b.size)
			LogFile.WriteLine(Now() & ":  Logout - " & loggedinUser &  ": LP " & userRoleFileLM & " size = " & c.size)
			Wscript.Quit
		End If

		If(userRoleFileAtShareLM<userRoleFileLM) Then 
			If fso.FileExists(userRoleFileAtShare) then
				fso.DeleteFile userRoleFileAtShare
			End If
			Call fso.CopyFile (("C:\ProgramData\Launchpad\" & userRoleFile), (userRoleFileAtShare),  True)
			If Err.Number <> 0 Then
				LogFile.WriteLine(Now() & ":  Logout - " & loggedinUser & ": Error while copying file from Launchpad folder to roleshare: " & Err.Description)
				Err.Clear
			End If	
			LogFile.WriteLine(Now() & ":  Logout - " & loggedinUser & ": File at share is old, copying from Launchpad folder")

		End If
		
	End If
	
	LogFile.Close
	Set LogFile = nothing
	
END SUB

Call Main()