Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim dt_Username, dtSidebarMenu, dtSidebarSubMenu, iteration

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKINV001-017 - Setup Mapping Jurnal Investasi Investasi - General Tambah, Ubah, View Detil & Hapus Data.xlsx", "DPLKINV001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, Array("Login Sebagai : " & dt_Username))

iteration = Environment.Value("ActionIteration")
REM ------- DPLK
'Call DA_Login()
'Call GoTo_SidebarMenu(dtSidebarMenu)
'Call GoTo_SidebarSubMenu(dtSidebarSubMenu)

'If iteration = 1 Then
'	Call AddSetupPortofolio()	
'ElseIf iteration = 2 Then
'	Call ViewSetupPortofolio()
'ElseIf iteration = 3 Then
'	Call EditSetupPortofolio()
'ElseIf iteration = 4 Then
'	Call DeleteSetupPortofolio()
'End If
'
'Call DA_Logout("0")

Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDPLK, LibReport, LibRepo, objSysInfo
	Dim tempDPLKPath, tempDPLKPath2, PathDPLK
	
	Set objSysInfo 		= Createobject("Wscript.Network")	
	
	tempDPLKPath 	= Environment.Value("TestDir")
	tempDPLKPath2 	= InStrRev(tempDPLKPath, "\")
	PathDPLK 		= Left(tempDPLKPath, tempDPLKPath2)
	
	LibPathDPLK	= PathDPLK & "Lib_Repo_Excel\LibDPLK\"
	LibReport			= PathDPLK & "Lib_Repo_Excel\LibReport\"
	LibRepo				= PathDPLK & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- DPLK lib
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Setup.qfl")
'	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
'	Call RepositoriesCollection.Add(LibRepo & "RP_Dashboard.tsr")
'	Call RepositoriesCollection.Add(LibRepo & "RP_Setup.tsr")
'	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	dt_Username					= DataTable.Value("USERID",dtLocalSheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
	REM --------- Menu
	dtSidebarMenu				= DataTable.Value("SIDEBAR_MENU", dtlocalsheet)
	dtSidebarSubMenu			= DataTable.Value("SIDEBAR_SUBMENU", dtlocalsheet)
End Sub
