﻿Dim dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult @@ script infofile_;_ZIP::ssf7.xml_;_
Dim preparation ,iteration

REM -------------- Call Function
Call spLoadLibrary()
Call spInitiateData("DPLKLib_Report.xlsx", "DPLKKPS001-010 - Kepesertaan - Setup Group Approval Tambah, Ubah, View Detil & Hapus Data.xlsx", "DPLKKPS001")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()

dtPreparation = Split(preparation, ";")
Call spAddScenario(dt_TCID, dt_TestScenarioDesc, dt_ScenarioDesc, dt_ExpectedResult, dtPreparation)
iteration = Environment.Value("ActionIteration")

REM ------- DPLK
Call DA_Login()
Call GoTo_SidebarMenu()
Call GoTo_SidebarSubMenu2()

If iteration = 1 Then
	Call AddSetupGroupApproval()
ElseIf iteration = 2 Then
	Call ViewSetupGroupApproval()
ElseIf iteration = 3 Then
	Call EditSetupGroupApproval()
ElseIf iteration = 4 Then
	Call DeleteSetupGroupApproval()
End If

Call ClickSidebarSubmenu()
Call ClickSidebar()
Call DA_Logout("0")

Call spReportSave()
	
Sub spLoadLibrary()
	Dim LibPathDPLK, LibReport, LibRepo, objSysInfo
	Dim tempDPLKPath, tempDPLKPath2, PathDPLK
	
	Set objSysInfo 	= Createobject("Wscript.Network")	
	
	tempDPLKPath 	= Environment.Value("TestDir")
	tempDPLKPath2 	= InStrRev(tempDPLKPath, "\")
	PathDPLK 		= Left(tempDPLKPath, tempDPLKPath2)
	
	LibPathDPLK		= PathDPLK & "Lib_Repo_Excel\LibDPLK\"
	LibReport		= PathDPLK & "Lib_Repo_Excel\LibReport\"
	LibRepo			= PathDPLK & "Lib_Repo_Excel\Repo\"

	REM ------- Report Library
	LoadFunctionLibrary (LibReport & "BNI_GlobalFunction.qfl")
	LoadFunctionLibrary (LibReport & "Run Report BNI.vbs")
	
	REM ---- DPLK lib
	LoadFunctionLibrary (LibPathDPLK & "DPLKLib_Menu.qfl")
	LoadFunctionLibrary (LibPathDPLK & "DPLK_Kepesertaan_Setup.qfl")
	Call RepositoriesCollection.Add(LibRepo & "RP_Login.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Administration_Dashboard.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Kepesertaan_Setup.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Sidebar.tsr")
	
End Sub

Sub spGetDatatable()
	REM --------- Data
	preparation		 			= DataTable.Value("PREPARATION",dtlocalsheet)
	
	REM --------- Reporting
	dt_TCID						= DataTable.Value("TC_ID", dtLocalSheet)
	dt_TestScenarioDesc			= DataTable.Value("TEST_SCENARIO_DESC", dtLocalSheet)
	dt_ScenarioDesc				= DataTable.Value("SCENARIO_DESC", dtLocalSheet)
	dt_ExpectedResult			= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
End Sub
