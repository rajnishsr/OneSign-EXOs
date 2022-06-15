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

	Dim rQLocation, rQLength, userRoleFile
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
				userRoleFile = RoleFileName 
			End If
		End IF 
	next
	ReadSSOGroups=userRoleFile
    
END FUNCTION

FUNCTION ReadRoleFiletoFindMaster(userRoleFile) 
	'Read Role.ini and find out the Master File.
     DIM objFileToRead, path, strLine, strParameter, fso, split0, split1, Master, MasterFile
	 SET fso=CreateObject("Scripting.FileSystemObject")
	 MasterFile =""
	 
	 path= "C:\ProgramData\Launchpad\" & userRoleFile

	 Set objFileToRead = fso.OpenTextFile(path, 1, False)

	 do While objFileToRead.AtEndOfStream = False
		strLine = Trim(objFileToRead.ReadLine)
	'Go to SSOIniTool and then iterate line by line
	
		If UCase(strLine)="[SSOINITOOL]" Then

			strLine = Trim(objFileToRead.ReadLine)
			Do While (Left(strLine,1)<> "[")
				strParameter = Split(strLine, "=")
				split0 = Trim(strParameter(0))
				split1 = Trim(strParameter(1))

				If UCase(split0) = "ARGUMENTS" THEN
					Master = split1				
					Exit Do
				End If
				
				strLine = Trim(objFileToRead.ReadLine)
			Loop	 
			strParameter = Split (Master, "AppParams")
			Master = replace(Trim(strParameter(0)), chr(34), "")
			If Right(Master,1) = " " Then
				Master = Left(Master, Len(Master)-1)
			End If
			
			If Right(Master,2) = "\\" Then
				Master = Left(Master, Len(Master)-1)
			End If
			
			If InStr(Master,".ini")>0 Then
				MasterFile = Master
			Else
				MasterFile = Master & "MasterWin7.ini"
			End If
			Exit Do
		End If
	 loop
		objFileToRead.Close
		
	ReadRoleFiletoFindMaster = MasterFile
END FUNCTION

FUNCTION ReadRoleFileTillTABS(userRoleFile) 
	'Read the whole Role.ini File.
     DIM objFileToRead, path, fso, comment
	 SET fso=CreateObject("Scripting.FileSystemObject")
	 comment =""
	 
	 path= "C:\ProgramData\Launchpad\" & userRoleFile

	 Set objFileToRead = fso.OpenTextFile(path, 1, False)
	 comment = objFileToRead.ReadAll()

	objFileToRead.Close
	ReadRoleFileTillTABS = comment
END FUNCTION

FUNCTION ReadRoleFileTillSnippets(userRoleFile) 
	'Read Role.ini and find out the user selection of buttons.
     DIM path, strLine, fso, Master, nonsnippet, objStream, adReadLine
	 
	 Set objStream = CreateObject("ADODB.Stream")
	 nonsnippet =""
	 path= "C:\ProgramData\Launchpad\" & userRoleFile
	 adReadLine = -2
	 objStream.Open
	 objStream.CharSet = "_autodetect_all"
	 
	 objStream.LoadFromFile path

	 strLine = objStream.ReadText(adReadLine)
 
	 do Until UCase(strLine)="[TABS]"
		strLine = Trim(objStream.ReadText(adReadLine))
	 loop

	 do 
		nonsnippet = nonsnippet & strLine & vbCrLf
		strLine = Trim(objStream.ReadText(adReadLine))

	 loop Until InStr(strLine, "[") >0
		objStream.Close
	ReadRoleFileTillSnippets = nonsnippet
END FUNCTION

FUNCTION ReadMasterFile(masterFile) 
	'Read Master File to find all the Snippets
     DIM objStream, objFileToRead, path, strLine, fso, buffer
	 Set objStream = CreateObject("ADODB.Stream")
	 path= masterFile
	 adReadLine = -2
	 objStream.Open
	 objStream.CharSet = "_autodetect_all"
	 buffer= ""
	 objStream.LoadFromFile path
	 
	 strLine = objStream.ReadText(adReadLine)
	 do Until UCase(strLine)="[TABS]"
		strLine = objStream.ReadText(adReadLine)
	 loop
		
	'First loop till TABS and then loop past the lines which start with [.
	'Once done, start storing the remaining in buffer
	 do 
		strLine = objStream.ReadText(adReadLine)

	 loop Until InStr(strLine, "[") >0
	 do 
		buffer=buffer & vbCrlF & strLine
		strLine = objStream.ReadText(adReadLine)
	 loop Until objStream.EOS
		ReadMasterFile = buffer 
		objStream.Close
END FUNCTION

Sub Main

	DIM objStream, masterFile, LogFile, fso, loggedinUser, roleQuery, userRole, userRoleFile, userRoleFile_copy, userRoleFilePath, Log, LogFilePath
	DIM doLPRefresh, masterFileName, masterFileLM, roleFileLM, buffer
	
	SET fso=CreateObject("Scripting.FileSystemObject")
	
	LogFilePath = "C:\ProgramData\Launchpad\log\UpdateRoleFile_Log.txt"

'Checking to see if log file exists, if not, create it.
'If log file is more than 50 MB, delete and recreate.

	If fso.FileExists(LogFilePath) Then
		Set Log = fso.GetFile(LogFilePath)
		If Log.Size > 500000 Then
			Log.Delete(True)
			Set LogFile = fso.CreateTextFile(LogFilePath, true)
		Else
			Set LogFile = fso.OpenTextFile(LogFilePath, 8, True, 0)
		End If
	Else
		Set LogFile = fso.CreateTextFile(LogFilePath, true)
	End If
	
	LogFile.Write "---------------------------------------------------------------------------------" & vbCrLf
	LogFile.WriteLine "Current Time = " & Now()
	'loggedinUser = "FJO9655"
	loggedinUser ="{VAR SSOUSR}"
	LogFile.WriteLine "Logged In User = " & loggedinUser
	
	'Find User Rolequery from Launchpad.ini
	roleQuery= ReadLaunchpadIniToFindRoleQuery()
	LogFile.WriteLine "roleQuery = " & roleQuery
	
	'Using RoleQuery amd loggedinuser, find the role file name associated with user.
	userRole= ReadSSOGroups(loggedinUser, roleQuery)
	LogFile.WriteLine "userRole = " & userRole
	
	userRoleFile = userRole & ".ini"
	'If for some reason, the user role file comes as blank, quit, else assign the file
	IF userRole="" Then
		LogFile.WriteLine "For some reason User Role File came as Blank, quitting."
		LogFile.WriteLine "Try checking if the Launchpad.ini file is in utf-8 format."
		Wscript.Quit
	Else 
		userRoleFilePath = "C:\ProgramData\Launchpad\" & userRoleFile 
		If fso.FileExists(userRoleFilePath) Then
			LogFile.WriteLine "User role file found in Launchpad folder."
		Else
			LogFile.WriteLine "User role file not found in Launchpad folder, quitting."
			Wscript.Quit
		End If
	END IF
		
	masterFileName = ReadRoleFiletoFindMaster(userRoleFile)
	LogFile.WriteLine "Master File = " & masterFileName
	
	Set masterFile =fso.GetFile(masterFileName)
	
   'Check if Button Manager is running. 
   'No need to update role file if Button Manager is running as it will eventually update the role file.
   Set objProcs=GetObject("winmgmts:\\.\root\cimv2").ExecQuery("select * from Win32_Process where Name= 'ButtonManager.exe'")
   For Each process In objProcs
      LogFile.WriteLine "Looks like Button Manager is running at this time, quitting."
	  Wscript.Quit
      On Error Resume Next
   Next

	IF fso.FileExists(masterFileName) Then
		masterFileLM = masterFile.DateLastModified
		LogFile.WriteLine "Master File Last Modified Date/ Time = " & masterFileLM
	ELSE
		LogFile.WriteLine "For some reason Master File came as Blank, quitting."
		LogFile.WriteLine "Try checking if the role file is in utf-8 format."
		Wscript.Quit
	End IF	
	
	'Copy the whole role file.
	roleFileLM = ReadRoleFileTillTABS(userRoleFile) 
	
	'In the role file, find if the Master file last modified date/time is in there.
	'If it is there then no need to update as the role file has the latest snippets.
	
	If InStr(roleFileLM, masterFileLM)>0 Then
		LogFile.WriteLine "There has been no change in Master File since last change, quitting."
		Wscript.Quit
	Else 
		userRoleFile_copy = "C:\ProgramData\Launchpad\" & userRole & "_copy.ini"
		LogFile.WriteLine "userRoleFile_copy = " & userRoleFile_copy
		
		'Check if a file with name userRoleFile_copy.ini is in the Launchpad folder, if it is there, delete it.
		If fso.FileExists(userRoleFile_copy) Then
			fso.DeleteFile(userRoleFile_copy)
			LogFile.WriteLine "Old _copy File deleted." 
		End If
		
		'Copy to create a role dile named userRoleFile_copy.ini in the Launchpad folder
		Call fso.CopyFile (userRoleFilePath, userRoleFile_copy, True)

		Set objStream = CreateObject("ADODB.Stream")
		
		'Writing the contecnt to be written to userRoleFile_copy.ini to a buffer and then updating the userRoleFile_copy.ini with the buffer
		buffer = "# Master File Last Modified : " & masterFileLM & vbCrLf
		buffer = buffer & ReadRoleFileTillSnippets(userRoleFile)
		buffer = buffer & ReadMasterFile(masterFile)
		
		path= userRoleFile_copy
		objStream.Open
		objStream.LoadFromFile path
		objStream.CharSet = "utf-8"
        objStream.WriteText buffer
		objStream.SaveToFile path, 2
		
		LogFile.WriteLine "Writing to Role File copy" 
		fso.DeleteFile userRoleFilePath
		LogFile.WriteLine "Delete old Role File" 
		fso.MoveFile userRoleFile_copy, userRoleFilePath
		LogFile.WriteLine "Renaming userRoleFile_copy.ini to userRoleFile.ini" 
	 End If
	 LogFile.WriteLine "Finishing Up."
	 LogFile.Close
	 Set LogFile = nothing

END SUB

Call Main()