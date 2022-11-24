Call ExecuteSQL()
REM-----Fungsi Mengeksekusi Query
Sub ExecuteSQL
	Set objkey = CreateObject("WScript.Shell")
	Dim query1
		query1		= DataTable.Value("QUERY_SETUP_PORTOFOLIO", "DPLKDBConfig")
		query1 		= Replace(query1, "<OBJECT_1>", Parameter("Jurnal"))
		Expl_Query1	= DataTable.Value("EXPL_QUERY1", "DPLKDBConfig")		
	wait 2
	objkey.SendKeys "{DOWN 3}"
	wait 1
	'buka tab field query baru
		objkey.SendKeys("^t")
	wait 3
		Window("HeidiSQL 12.0.0.6468").WinObject("Query_2").Click
	'Mengisi Field query dengan query dari excel
	wait 3
		objkey.SendKeys("^a")
		objkey.SendKeys "{BACKSPACE}"
		If len(DataTable.Value("QUERY_SETUP_PORTOFOLIO", "DPLKDBConfig")) > 600  Then
		wait 10
			objkey.SendKeys(query1)
		wait 20
		else
		objkey.SendKeys(query1)
		End If
		
	'Running query	
		wait 10
		objkey.SendKeys "{F9}"
		wait 10
	Call CaptureImageUFTV2(Window("HeidiSQL 12.0.0.6468"), Expl_Query1 , " ", compatibilityMode.Desktop, ReportStatus.Passed)		
		
	'Menutup field query
'	Window("HeidiSQL 12.0.0.6468").WinObject("Query_2").Click
	wait 4
		objkey.SendKeys "^{F4}"
	wait 2
	
End Sub
