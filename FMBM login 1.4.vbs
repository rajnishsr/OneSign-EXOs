'FMBM Login Version 1.4 date 03/08/2022
'Added functionality to check if the role file at share or launchpad is not editable. If they are not, code quits.
'Added functionality to check at the file size at share and Launchpad and if one of them is 0 KB, code quits
'Using Lp.exe /role to tell Launchpad which role file to use. This will remove the unnecessary AD call.
'Added functionality to check for previous logged in user and if it is different, copy the old role file back.



FUNCTION ReadLaunchpadIniToFindRoleQuery() 
	'Read Launchpad.ini and find out the Role Query used.
     DIM path, strLine, strParameter, fso, split0, split1
	 SET fso=CreateObject("Scripting.FileSystemObject")
	 path= "C:\ProgramData\Launchpad\LaunchPad.ini"
	 Set objFileToRead = fso.OpenTextFile(path, 1, False, -1)

	 do Until objFileToRead.AtEndOfStream
		strLine = objFileToRead.ReadLine()
		If InStr(strLine, "=")>0 Then
			strParameter = Split(strLine, "=")
			split0 = Trim(strParameter(0))
			split1 = Trim(strParameter(1))
			IF UCase(split0) = "ROLEQUERY" THEN
				roleQuery = split1
			END IF
		End If
	 loop
		objFileToRead.Close
		
	ReadLaunchpadIniToFindRoleQuery	= roleQuery
END FUNCTION	
   
FUNCTION ReadSSOGroups(userName, roleQuery)

	Dim rQLocation, rQLength, userRole, userRoleFile
	Set objNetwork = WScript.CreateObject("WScript.Network")
	Set FSO = CreateObject("Scripting.FileSystemObject")

	domain = "DC=hca,DC=corpad,DC=net"
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	objConnection.Provider = ("ADsDSOObject")
	objConnection.Open "Active Directory Provider"
	objCommand.ActiveConnection = objConnection

	objCommand.CommandText = "<LDAP://DC=hca,DC=corpad,DC=net>;(samAccountName=" & username & ");memberOf;subTree"
	Set objRecordSet = objCommand.Execute
	result = objRecordSet.Fields("memberOf").Value

	rQLength = Len(roleQuery)
	For Each objGroup In result

		strGroupName=objGroup
		rQueryLen = InStr(strGroupName, ",")

		rQLocation= InStr(UCase(strGroupName), UCase(roleQuery))
		IF rQLocation>0 Then
			
			RoleFileName = Mid(strGroupName, (rQLocation+rQLength), (rQueryLen-rQLocation-rQLength))
			roleFile = "C:\ProgramData\Launchpad\" & RoleFileName & ".ini"
			If FSO.FileExists(roleFile) Then
				userRoleFile = RoleFileName & ".ini"
			End If
		End IF 
	next
	ReadSSOGroups=userRoleFile
    
END FUNCTION

FUNCTION IsWriteAccessible(sFilePath)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
	IsWriteAccessible = False

    Dim fso 
	Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next

    Dim nErr : nErr = 0
    Dim sDesc : sDesc = ""
    Dim oFile : Set oFile = fso.OpenTextFile(sFilePath, ForAppending)
    If Err.Number = 0 Then
        oFile.Close
        If Err Then
            nErr = Err.Number
            sDesc = Err.Description
        Else
            IsWriteAccessible = True
        End if
    End If
	
End FUNCTION

Sub Main

	DIM objShell, objFileToRead, strFolder, loggedinUserCmd, FMBMFolder, FMBMFile, strLine, iCounter
	DIM userRoleFileAtShare, userRoleFileAtShareLM, userRoleFileLM, a, b, c, LogFile, LPRefreshCmd, TypeKey, FMBMLog, FMBMFileSize, LPRoleFile, FMBMLoggedInUser, FMBMLoggedInUserFile
	DIM fso, loggedinUser, prevloggedinUser, UserComp, shareLocation, roleQuery, userRole, userRoleFile, AgentType, userRoleFileFMBM, FMBMLoggedInUserFile1
	DIM ErrorBuffer, doLPRefresh, LPRoleFileDLM, userRoleFileFMBMDLM, FMBFldr, FMBMFiles
	
	SET fso=CreateObject("Scripting.FileSystemObject")
	SET objShell = CreateObject("WScript.Shell")	
	On Error Resume Next
	
	'Error Buffer helps in collecting error messages till Error file is created.
	'Once FMBMLog is created, Error Buffer will write to that file
	ErrorBuffer ="---------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf
	
	'loggedinUser = "FJO9655"
	loggedinUser ="{VAR SSOUSR}"
	ErrorBuffer = ErrorBuffer & Now() & ":  Login - " & loggedinUser & " <--Logged in User"  & vbCrLf
	
	'Create folders if they do not exist	
	IF Not fso.FolderExists("C:\ProgramData\Launchpad\INIToolBackup\") Then
		Set FMBMfolder=fso.CreateFolder("C:\ProgramData\Launchpad\INIToolBackup\")
		ErrorBuffer = Now() & ":  Login - " & loggedinUser & ": INIToolBackup folder created." & vbCrLf
		Set FMBMfolder =  nothing
	END IF
		
	If Err.Number <> 0 Then
		ErrorBuffer = Now() & ":  Login - " & loggedinUser & ": Error while creating INIToolBackup folder: " & Err.Description & vbCrLf
		Err.Clear
	End If		
	
	IF Not fso.FolderExists("C:\ProgramData\Launchpad\INIToolBackup\FMBM") Then
		Set FMBMfolder=fso.CreateFolder("C:\ProgramData\Launchpad\INIToolBackup\FMBM")
		ErrorBuffer = Now() & ":  Login - " & loggedinUser & ": FMBM folder created." & vbCrLf
		Set FMBMfolder =  nothing
	END IF
		
	If Err.Number <> 0 Then
		ErrorBuffer = Now() & ":  Login - " & loggedinUser & ": Error while creating FMBM folder: " & Err.Description & vbCrLf
		Err.Clear
	End If		
	

	IF Not fso.FolderExists("C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\") Then
		Set FMBMfolder=fso.CreateFolder("C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\")
		ErrorBuffer = Now() & ":  Login - " & loggedinUser & ": FMBMBkUp folder created." & vbCrLf
		Set FMBMfolder =  nothing
	END IF	
		
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Login - " & loggedinUser & ": Error while creating FMBM Backup folder: " & Err.Description & vbCrLf
		Err.Clear
	End If
	
	FMBMLoggedInUser = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\LoggedInUser.txt"
	prevloggedinUser = "NoPrevUser"
	'prevloggedinUser = "fjo9656"
	
	
	If fso.FileExists(FMBMLoggedInUser) Then
		'Set FMBMLoggedInUserFile = fso.GetFile(FMBMLoggedInUser)
		Set FMBMLoggedInUserFile = fso.OpenTextFile(FMBMLoggedInUser,1, false, -2)

		do Until FMBMLoggedInUserFile.AtEndOfStream
			strLine = Trim(FMBMLoggedInUserFile.ReadLine)
			
			If Left(strLine, 1) =";" or Left(strLine, 1) = "#" Then
				'Yep, this string is commented!
			Else 
				If Not IsNull(strLine) Then
					prevloggedinUser =  strLine
				End If
			End If
		loop
		FMBMLoggedInUserFile.Close()
		Set FMBMLoggedInUserFile = nothing
		
		fso.DeleteFile(FMBMLoggedInUserFile)
	End If
	
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Login - " & loggedinUser & ": Checked for loggedinUser file" & vbCrLf
		Err.Clear
	End If
	'Delete the LoggedinUserFile
	
	Set FMBMLoggedInUserFile1 = fso.CreateTextFile(FMBMLoggedInUser, True)
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Login - " & loggedinUser & ": 1: " & Err.Description & vbCrLf
		Err.Clear
	End If

	FMBMLoggedInUserFile1.Write loggedinUser
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Login - " & loggedinUser & ": 2: " & Err.Description & vbCrLf
		Err.Clear
	End If

	FMBMLoggedInUserFile1.Close
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Login - " & loggedinUser & ": 3: " & Err.Description & vbCrLf
		Err.Clear
	End If

	Set FMBMLoggedInUserFile1 = nothing
	
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Login - " & loggedinUser & ": Error while deleting/ creating loggedinUser file: " & Err.Description & vbCrLf
		Err.Clear
	End If

	FMBMFile = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\" &"FMBM_" & loggedinUser & ".txt"
	FMBMLog = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMLog.txt"	
	
	If fso.FileExists(FMBMLog) Then
		Set FMBMFileSize = fso.GetFile(FMBMLog)
		If FMBMFileSize.Size > 5000000 Then
			FMBMFileSize.Delete(True)
		End If
		Set LogFile = fso.OpenTextFile(FMBMLog, 8, True, 0)
		'Set LogFile = fso.GetFile(FMBMLog)
	Else
		Set LogFile = fso.CreateTextFile(FMBMLog, true)
	End If
	
	If Err.Number <> 0 Then
		ErrorBuffer= ErrorBuffer & Now() & ":  Login - " & loggedinUser & ": Error while checking if FMBMLog file exists: " & Err.Description & vbCrLf
		 Err.Clear
	End If
	'Write whatever is in ErrorBuffer
	LogFile.Write ErrorBuffer
	
	If Err.Number <> 0 Then
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Just after dumping the Buffer " & Err.Description)
		Err.Clear
	End If
		
	LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": FMBM file=  " & FMBMFile)
	'Check if the FMBM file for the user exists. This file has all the config for the user.
	If Err.Number <> 0 Then
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Before searching for FMBM file: " & Err.Description)
		Err.Clear
	End If
	
	If fso.FileExists(FMBMFile) Then
	'If file exists read shareLocation, roleQuery and userRoleFile from the file
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": User property file exists, checking for values")
		Set objFileToRead = fso.OpenTextFile(FMBMFile,1)
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


		'Find User Rolequery from Launchpad.ini
		roleQuery= ReadLaunchpadIniToFindRoleQuery()
		'Using RoleQuery amd loggedinuser, find the role file associated with user.
		userRoleFile= ReadSSOGroups(loggedinUser, roleQuery)
		'If for some reason, the user role file comes as blank, quit
		IF userRoleFile="" Then
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Exiting code as userRoleFile is blank")
			Wscript.Quit
		End IF	
		
		'Find share location	
		If fso.FileExists("C:\ProgramData\Launchpad\Roleshare.ini") Then

			Set objFileToRead = fso.OpenTextFile("C:\ProgramData\Launchpad\Roleshare.ini",1)
			do while not objFileToRead.AtEndOfStream
				strLine = objFileToRead.ReadLine()
				strLine = Trim(strLine)
				IF Right(strLine, 1) <> "\" Then
					strLine = strLine & "\"
				End If
				If Left(strLine, 1) =";" or Left(strLine, 1) = "#" Then
					'Yep, this string is commented!
				Else
					IF fso.FolderExists(strLine) Then
						shareLocation=strLine
					End IF
				End If
			loop
			
			objFileToRead.Close
			Set objFileToRead = Nothing

			Else
				'If the roleshare location file do not exist, exit code.
				LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Exiting code as Roleshare.ini do not exist at C:\ProgramData\Launchpad")
				Wscript.Quit
		End If
		
		If Err.Number <> 0 Then
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Share Location is not accessible: " & Err.Description)
			Err.Clear
		End If		
			
		'If FMBM file don't exist, find shareLocation, roleQuery, userRoleFile and agent type values and enter in FMBM_User file	
		If  fso.FolderExists(shareLocation) Then
			Set a = fso.CreateTextFile(FMBMFile, true)
			a.WriteLine(shareLocation)
			a.WriteLine(roleQuery)
			a.writeLine(userRoleFile)

			TypeKey = "HKLM\SOFTWARE\WOW6432Node\SSOProvider\ISXAgent\Type"
			AgentType = objShell.regread(TypeKey)
			a.writeLine(AgentType)
			a.Close
			Set a = nothing
		Else
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": FMBM file not created as share is not accessible")
		End If
	End If
	
	If Err.Number <> 0 Then
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Cleaning Error buffer: " & Err.Description)
		Err.Clear
	End If	
	
	LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": shareLocation = " & shareLocation)
	LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": roleQuery = " & roleQuery)
	LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": userRoleFile = " & userRoleFile)
	LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": AgentType = " & AgentType)
	
	'If share is not accessible, quit
	If Not fso.FolderExists(shareLocation) Then
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Exiting code as share is not accessible")
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Exiting code as share is not accessible.")
		Wscript.Quit
	End If
	
	doLPRefresh = 0
	'If LoggedinUser is different then Previous Logged in User, change the role file
	Set FMBFldr = fso.GetFolder("C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\"	)
	Set FMBMFiles = FMBFldr.Files
	If FMBFldr.Files.Count <>0 Then
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": FMBMBkUp is not empty")
		For Each item in FMBMFiles
			If UCase(fso.GetExtensionName(item.name)) = "INI" Then
				userRoleFileFMBM = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\" & item.Name
				LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": FMBM BkUp File is " & userRoleFileFMBM)
			End If
		Next
	End If
			 
	If Err.Number <> 0 Then
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Error while finding FMBM Backup file " & Err.Description)
		Err.Clear
	End If	
	
	'userRoleFileFMBM = "C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\" & userRoleFile
	UserComp = StrComp(loggedinUser, prevloggedinUser, 1)
	
	LogFile.WriteLine(Now() & ": Login - " & loggedinUser & ": Number of Files in FMBMBkUp are " & FMBMFldr.Files.Count)
	LPRoleFile = "C:\ProgramData\Launchpad\" & userRoleFile
	
	If FMBFldr.Files.Count = 0 Then
		'Copy current role file to FMBM folder
		Call fso.CopyFile (LPRoleFile, "C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\", True)
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Copying Role file from Launchpad to FMBMBkUp")
		
		If Err.Number <> 0 Then
			'LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Error while copying role file from Launchpad to FMBM Backup folder : " & Err.Description)
			Err.Clear
		End If		
	Else
		If UserComp <> 0 and prevloggedinUser <> "NoPrevUser" Then
			LogFile.WriteLine(Now() &  ":  Login - " & prevloggedinUser & " <--Previous Logged in User")
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Logged in User is different from Previous User")
			 If fso.FileExists(userRoleFileFMBM) Then
				'First copy FMBM file to Launchpad
				Call fso.CopyFile ((userRoleFileFMBM),("C:\ProgramData\Launchpad\"), True)
				fso.DeleteFile("C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\*.*"), DeleteReadOnly
				'Then copy current role file to FMBM folder
				Call fso.CopyFile (LPRoleFile, "C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBMBkUp\", True)
				doLPRefresh =1
				LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Copying role file from FMBM Backup folder back to Launchpad.")	
			 End If
			 If Err.Number <> 0 Then
				LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Finished copying previous role file from FMBM Backup folder to Launchpad and vice versa" )
				Err.Clear
			 End If
		Else
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Either the Logged in User is same as previous User or there is no previous user")
		End If
	End If

	If Err.Number <> 0 Then
		'LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Error while copying current role file to FMBMBkUp folder: " & Err.Description)
		Err.Clear
	End If
	
	userRoleFileAtShare = shareLocation & loggedinUser &"_" & userRoleFile
	
	'If User Role file exists at share, compare it with the Launchpad folder.
	'Whichever is newer is copied to the other location.
	'If the fie does not exist, copy the file from Launchpad folder to Role share folder.
	
	Err.Clear
	If fso.FileExists(userRoleFileAtShare) Then 
		Set b = fso.GetFile(userRoleFileAtShare)
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": User Role file at share exists and is " & userRoleFileAtShare)
		userRoleFileAtShareLM = b.DateLastModified
		Set c = fso.GetFile("C:\ProgramData\Launchpad\" & userRoleFile)
		userRoleFileLM = c.DateLastModified
		
		'Check to see if the role file at share or local is 0 KB, if it is then quit
		If(b.size = 0 or c.size = 0) Then
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Exiting code as file size is 0")
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Share " & userRoleFileAtShareLM & " size = " & b.size)
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": LP " & userRoleFileLM & " size = " & c.size)
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Exiting code as Role File is 0 KB")
			Wscript.Quit
		End If
		
		'Check to see if the role file at share or local can be modified, if not then quit
		If(Not(IsWriteAccessible(b)) or Not(IsWriteAccessible(c))) Then
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Exiting code as file cannot be modified.")
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": File at Share " & userRoleFileAtShare & " is editable - " & IsWriteAccessible(b))
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": File at C:\ProgramData\Launchpad\" & userRoleFile & " is editable - " & IsWriteAccessible(c))
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Exiting code as Role File is not editable.")
			Wscript.Quit
		End If
				
		
		If(userRoleFileAtShareLM>userRoleFileLM) Then
			Call fso.CopyFile ((userRoleFileAtShare), ("C:\ProgramData\Launchpad\" & userRoleFile), True)
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": File at share Last Modified Date is " & userRoleFileAtShareLM)
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Last Modified Date of Role File " & userRoleFile & " at C:\ProgramData\Launchpad\ is " & userRoleFileLM)
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": File at share is newer, copying it to Launchpad folder")
			Set objFileToRead = fso.OpenTextFile("C:\ProgramData\Launchpad\" & userRoleFile, 1, False, -1)
			objFileToRead.close
			set objFileToRead = nothing
			'Setting the Launchpad to be refreshed
			doLPRefresh = 1
			
			If Err.Number <> 0 Then
				LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Error while copying file at share to Launchpad folder : " & Err.Description)
				Err.Clear
			End If
		End If
		If(userRoleFileAtShareLM<userRoleFileLM) Then 
			If fso.FileExists(userRoleFileAtShare) then
				fso.DeleteFile userRoleFileAtShare
			End If
			Call fso.CopyFile (("C:\ProgramData\Launchpad\" & userRoleFile), (userRoleFileAtShare),  True)
		
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": File at share Last Modified Date is " & userRoleFileAtShareLM)
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser &  ": Last Modified Date of Role File " & userRoleFile & " at C:\ProgramData\Launchpad\ is " & userRoleFileLM)
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": File at share is old, copying from Launchpad folder")
			
			Set objFileToRead = fso.OpenTextFile(userRoleFileAtShare, 1, False, -1)
			objFileToRead.close
			set objFileToRead = nothing			
			
			If Err.Number <> 0 Then
				LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Error while copying file from Launchpad folder to roleshare: " & Err.Description)
				Err.Clear
			End If	
			

		End If
	Else
		Call fso.CopyFile (("C:\ProgramData\Launchpad\" & userRoleFile), (userRoleFileAtShare),  True)

		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": No file at share, copying from Launchpad folder")
		
		If Err.Number <> 0 Then
			LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Error while copying from Launchpad folder to share: " & Err.Description)
			Err.Clear
		End If
		Set objFileToRead = fso.OpenTextFile(userRoleFileAtShare, 1, False, -1)
		objFileToRead.close
		set objFileToRead = nothing
	End If
	
	'If the task has been set to refesh Launchpad, do it now
	If doLPRefresh = 1 Then
		iCounter=InStr(userRoleFile,".ini")
		RoleFileName= Left(userRoleFile, iCounter-1)
		objShell.run "Taskkill /F /IM ""LP.exe""", 1 , true
		objShell.run("""C:\Program Files (x86)\Imprivata\OneSign Agent\x64\LP.exe""" & " /role " & RoleFileName)
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Launchpad has been refreshed with " & RoleFileName)
	End If
	If Err.Number <> 0 Then
		LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Error while refreshing Launchpad: " & Err.Description)
		Err.Clear
	End If	
	LogFile.WriteLine(Now() & ":  Login - " & loggedinUser & ": Finishing Login script")
	LogFile.Close
	Set LogFile = nothing
END SUB

Call Main()