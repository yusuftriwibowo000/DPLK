Call CreateSessionHeidi()

REM-----Create Session Heidi tanpa capture ke report
Sub CreateSessionHeidi
	Call spCloseProgram()
	wait 5
	Call spOpenProgram()
	wait 5
	If Window("Check for HeidiSQL updates").Exist(5) Then
		Window("Check for HeidiSQL updates").WinObject("Skip").Click
	End If
	wait 2
	Window("HeidiApps").WinObject("NewSession").Click
	wait 1
	
	Dim dtUser, dtPassword, dtHostname
	dtUser 		= DataTable.Value("USER_DB", "DPLKDBConfig")
	dtPassword	= DataTable.Value("PASSWORD_DB", "DPLKDBConfig")
	dtHostname	= DataTable.Value("HOSTNAME", "DPLKDBConfig")
	
	Window("HeidiApps").WinComboBox("ComboBoxNetworkType").Click
	Window("HeidiApps").WinComboBox("ComboBoxNetworkType").Click

	Set objkey = CreateObject("WScript.Shell")
	wait 1
			objkey.SendKeys "{DOWN 5}"
	wait 1
			objkey.SendKeys "{ENTER}"
	wait 1
			objkey.SendKeys "{TAB 2}"
	wait 1			
			objkey.SendKeys(dtHostname)
	wait 1		
			objkey.SendKeys "{TAB 3}"
	wait 1		
			objkey.SendKeys(dtUser)
	wait 1		
			objkey.SendKeys "{TAB}"
	wait 1		
			objkey.SendKeys(dtPassword)
	wait 5
			objkey.SendKeys "{ENTER}"
	wait 10
			objkey.SendKeys "{ENTER}"
	wait 5
			Window("HeidiSQL 12.0.0.6468").Maximize
	wait 2
End Sub

'Fungsi buka aplikasi
Function spOpenProgram()
	OpenProgram1		= DataTable.Value ("PROGRAM1", dtGlobalSheet)
	program_path1		= DataTable.Value ("PROGRAM_PATH1", dtGlobalSheet)
	InvokeApplication(program_path1 & "/" & openprogram1 & ".exe")
End Function

'Fungsi tutup aplikasi
Function spCloseProgram()
	CloseProgram		= DataTable.Value ("PROGRAM1", dtGlobalSheet)
	SystemUtil.CloseProcessByName CloseProgram &".exe"
End Function

Sub loadLibrary
	Dim PathDPLk, LibRepo
	PathDPLK	= Environment.Value("Path_Folder")
	LibRepo		= PathDPLK & "Lib_Repo_Excel\Repo\"
	Call RepositoriesCollection.Add(LibRepo & "RP_Heidi.tsr")
End Sub
