Call DA_SQLCheckpoint("MYSQL", Parameter("Jurnal"), Parameter("TransactionDate"), Parameter("NoIdentitas"))
Call DA_SQLCheckpoint("BILT", Parameter("Jurnal"), Parameter("TransactionDate"), Parameter("NoIdentitas"))
Call DA_SQLCheckpoint("GLIF", Parameter("Jurnal"), Parameter("TransactionDate"), Parameter("NoIdentitas"))

REM ------------------- VERIFIKASI DATABASE PAKE SQL DEVELOPER
Sub DA_SQLCheckpoint(ByVal QueryType, ByVal noJurnal, ByVal trxDate, ByVal dtSQL_Query_AGN0181)
	Dim iCounter
	Dim iMaxLimit
	Dim dt_TCID

	dt_TCID			= Parameter("TCID")

	iMaxLimit = 300
	
	If QueryType = "BILT" Then
		dtConnection 	= DataTable.Value("ORACLE_CONNECTION", "KeagenanDBConfig")
		'dtUser			= DataTable.Value("ORACLE_USER", "KeagenanDBConfig")
		dtPswd			= DataTable.Value("ORACLE_PASSWORD", "KeagenanDBConfig")
		dtSQL_Query 	= DataTable.Value("BILT_QUERY", "KeagenanDBConfig")
		dtSQL_Query		= Replace(dtSQL_Query, "<JURNAL_NUMBER>", noJurnal)
		
	ElseIf QueryType = "GLIF" Then	
		dtConnection 	= DataTable.Value("ORACLE_CONNECTION", "KeagenanDBConfig")
		'dtUser			= DataTable.Value("ORACLE_USER", "KeagenanDBConfig")
		dtPswd			= DataTable.Value("ORACLE_PASSWORD", "KeagenanDBConfig")
		dtSQL_Query 	= DataTable.Value("GLIF_QUERY", "KeagenanDBConfig")
		dtSQL_Query		= Replace(dtSQL_Query, "<JURNAL_NUMBER>", noJurnal)
		dtSQL_Query		= Replace(dtSQL_Query, "<TRANSACTION_DATE>", trxDate)
		
	ElseIf QueryType = "MYSQL" Then		
		dtConnection 	= DataTable.Value("MYSQL_CONNECTION", "KeagenanDBConfig")
		'dtUser			= DataTable.Value("MYSQL_USER", "KeagenanDBConfig")
		dtPswd 			= DataTable.Value("MYSQL_PASSWORD", "KeagenanDBConfig")
		dtSQL_Query 	= DataTable.Value("MYSQL_QUERY", "KeagenanDBConfig")
		dtSQL_Query		= Replace(dtSQL_Query, "<JURNAL_NUMBER>", noJurnal)
		
		dtSQL_Query_AGN0181 	= DataTable.Value("MYSQL_QUERY", "KeagenanDBConfig")
		dtSQL_Query_AGN0181		= Replace(dtSQL_Query_AGN0181, "<NO_IDENTITAS>", Parameter("NoIdentitas"))
			
	End If
	
	If not JavaWindow("Oracle SQL Developer").JavaToolbar("SQL Developer Toolbar").Exist(10) Then
		SystemUtil.Run "C:\sqldeveloper\sqldeveloper.exe"
		iCounter = 0
		Do
			iCounter = iCounter + 10
		Loop Until JavaWindow("Oracle SQL Developer").JavaToolbar("SQL Developer Toolbar").Exist(10) or iCounter > iMaxLimit
	End If
	
	JavaWindow("Oracle SQL Developer").Activate
	JavaWindow("Oracle SQL Developer").Maximize
	
	JavaWindow("Oracle SQL Developer").JavaToolbar("SQL Developer Toolbar").Press "sqlworksheet"
	wait 2
	JavaWindow("Oracle SQL Developer").JavaDialog("Select Connection").JavaList("Connection:").Select dtConnection
	JavaWindow("Oracle SQL Developer").JavaDialog("Select Connection").JavaButton("OK").Click
	JavaWindow("Oracle SQL Developer").JavaDialog("Connection Information").JavaEdit("Password:").Set dtPswd
	wait 2
	JavaWindow("Oracle SQL Developer").JavaDialog("Connection Information").JavaButton("OK").Click
	wait 2
	
	If dt_TCID = "AGN0181" Then
		JavaWindow("Oracle SQL Developer").JavaEdit("IdeEditorPane").Set dtSQL_Query_AGN0181
	Else
		JavaWindow("Oracle SQL Developer").JavaEdit("IdeEditorPane").Set dtSQL_Query
	End If
	wait 2
	JavaWindow("Oracle SQL Developer").JavaToolbar("Query Toolbar").Press "run"
	
	Dim isResultDisplayed
	
	iCounter = 0
	Do
		wait 5
		isResultDisplayed = JavaWindow("Oracle SQL Developer").JavaStaticText("Fetched Message").Object.isVisible()
		Wait 1
		iCounter = iCounter + 1
	Loop Until isResultDisplayed Or iCounter > iMaxLimit
	
	Judul_Verifikasi = "Verifikasi Database " & QueryType
	
	Call CaptureImageUFTV2(JavaWindow("Oracle SQL Developer"), Judul_Verifikasi, " ", compatibilityMode.Desktop, ReportStatus.Done)
	
	Set ws = CreateObject("wscript.shell")
	ws.sendkeys "^{F4}"
	JavaWindow("Oracle SQL Developer").JavaDialog("Save").JavaButton("No").Click
	
	If QueryType = "MYSQL" Then
		JavaWindow("Oracle SQL Developer").JavaTree("CustomTree").Select "Connections;MySQL Connections;MYSQL_KEAGENAN"
		JavaWindow("Oracle SQL Developer").JavaTree("CustomTree").Click 114,61,"RIGHT"
		JavaWindow("Oracle SQL Developer").JavaMenu("Disconnect").Select
	Else
		JavaWindow("Oracle SQL Developer").JavaTree("CustomTree").Select "Connections;Oracle Connections;ORACLE_KEAGENAN"
		JavaWindow("Oracle SQL Developer").JavaTree("CustomTree").Click 87,30,"RIGHT"
		JavaWindow("Oracle SQL Developer").JavaMenu("Disconnect").Select
	End If
	'SystemUtil.CloseProcessByName "sqldeveloper64W.exe"
End Sub
