Sub Main

	DIM fso
	SET fso=CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	
	IF fso.FolderExists("C:\ProgramData\Launchpad\INIToolBackup\FMBM") Then
		FMBMfolder=fso.DeleteFile("C:\ProgramData\Launchpad\INIToolBackup\FMBM\FMBM_*.txt")
	END IF
		
END Sub

Call Main()