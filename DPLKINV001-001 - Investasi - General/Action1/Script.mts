Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_Username, iteration

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKINV001-001 - Setup Portofolio Investasi - General Tambah, Ubah, View Detil & Hapus Data.xlsx", "DPLKINV001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_Username))

iteration = Environment.Value("ActionIteration")
REM ------- DPLK
Call DA_Login()
Call GoTo_SidebarMenu()
Call GoTo_SidebarSubMenu()

If iteration = 1 Then
	Call AddSetupPortofolio()
	Call DA_Logout("0")
ElseIf iteration = 2 or iteration = 5 or iteration = 7 Then	
	Call ViewLogAuditTrail()	
	Call DA_Logout("0")
	Call spVerification()
ElseIf iteration = 3 Then
	Call ViewSetupPortofolio()
	Call DA_Logout("0")
ElseIf iteration = 4 Then
	Call EditSetupPortofolio()
	Call DA_Logout("0")
ElseIf iteration = 6 Then
	Call DeleteSetupPortofolio()
	Call DA_Logout("0")
End If

Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDPLK, LibReport, LibRepo, objSysInfo
	Dim tempDPLKPath, tempDPLKPath2, PathDPLK
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
'	tempDPLKPath 	= Environment.Value("TestDir")
'	tempDPLKPath2 	= InStrRev(tempDPLKPath, "\")
'	PathDPLK 		= Left(tempDPLKPath, tempDPLKPath2)
	PathDPLK		= Environment.Value("Path_Folder")
	
	LibPathDPLK	= PathDPLK & "Lib_Repo_Excel\LibDPLK\"
	LibReport			= PathDPLK & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDPLK & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- DPLK lib
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Setup.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Laporan.qfl")
	LoadFunctionLibrary (LibPathDPLK & "Lib_Verifications.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Dashboard.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Setup.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Laporan.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_Username					= DataTable.Value("USERID",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
End Sub

Sub spVerification()
	Call spVerifikasi_Heidi_CreateSession()
	Call spVerifikasi_Heidi_Query("PM018")
End Sub
