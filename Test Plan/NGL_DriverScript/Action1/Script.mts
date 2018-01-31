Option Explicit
'*********************************************************
' DRIVER SCRIPT
'*********************************************************
'********************************************************
' Purpose:  Initialize the Driver object and run the test TestCaseID
' Inputs:   strFrameworkFolderPath path to the folder concaining the keyword driven framework
' Returns:  Nothing.
' Author: H
NewQTPReport
'NewDriver(strTestCasesFullPath, strTestDataFullPath)
'NewDriver
'*********************************************************

SystemUtil.CloseDescendentProcesses

'Variables
Dim strTestCasesFullPath
Dim strTestDataFullPath
Dim strTestCaseID,strTestCaseId1
Dim strtestBatchExecution
Dim strDependency
Dim strDependsOn
Dim strStatus
Dim intTestCaseCount1,m
'Dim strResultLogFullPath


'Parameter("TestCasesFullPath")="D:\Hitesh\Automation\NexGen Lite\Resources\Data\TestCases.xls"
'msgbox Parameter("TestCasesFullPath")

'Retrieve Parameters
strTestCasesFullPath = Parameter("TestCasesFullPath")
strTestDataFullPath = Parameter("TestDataFullPath")
strTestCaseID = Parameter("TestCaseID")
'strResultLogFullPath = Parameter("ResultLogFullPath")
'set myobject1 = CreateObject("XMLParser.XMLParser")
'Check if TestCaseID is empty
If strTestCaseID = "" Then
	MsgBox "The TestCaseID Parameter is empty.", vbOK, "Error: TestCaseID Parameter"
	ExitAction()
End If

'Set the Full Path to the TestCases and TestData Excel files as Exception in case of user forget to set the parameters
If strTestCasesFullPath = "" Then
	'strTestCasesFullPath = "..\Data\TestCases.xlsx"
	strTestCasesFullPath = "..\Data\TestCases.xls"
	'msgbox (PathFinder.Locate ("..\Data\TestCases.xls"))
	
End If
If strTestDataFullPath = "" Then
	strTestDataFullPath = "..\Data\TestData.xls"
End If

'Initiate Driver and run test
Dim aDriver,intTestBatchLoop, intTestCaseCount,strTestCaseIterations,intTestData,strTestIds
Set aDriver = NewDriver(strTestCasesFullPath, strTestDataFullPath)
intTestCaseCount = datatable.GetSheet("dTestBatchExecution").getrowcount
For intTestBatchLoop = 1 To intTestCaseCount
	datatable.GetSheet("dTestBatchExecution").SetCurrentRow(intTestBatchLoop)
	strTestCaseId = datatable("Test_ID", "dTestBatchExecution")	
	strDependency = datatable("Dependency", "dTestBatchExecution")                     
	strDependsOn = datatable("Depends_On", "dTestBatchExecution")	    
	intTestCaseCount1 = datatable.GetSheet("dTestBatchExecution").getrowcount
'	For m = 1 To intTestCaseCount1
	m=1
	strTestCaseIterations = datatable("Iterations", "dTestBatchExecution")
	strTestIds = datatable("Query", "dTestBatchExecution")	
	strTestIds=split(strTestIds,";")
	
	For intTestData = 1 To strTestCaseIterations
   	
		environment("TestDataID")= strTestIds(intTestData-1)
   		datatable.GetSheet("dTestBatchExecution").SetCurrentRow(m)
    	strTestCaseId1 = datatable("Test_ID", "dTestBatchExecution")
   		If Ucase(strDependency) = Ucase("Y") Then
 	
			If Ucase(strTestCaseId1) = Ucase(strDependsOn) Then
				strStatus = datatable("Status", "dTestBatchExecution")
		    	Select Case Ucase(strStatus)
		        	Case "PASSED"
		        		datatable.GetSheet("dTestBatchExecution").SetCurrentRow(intTestBatchLoop)
		        	    aDriver.RunTest(strTestCaseID) 
		        	Case "FAILED"
		        		datatable.GetSheet("dTestBatchExecution").SetCurrentRow(intTestBatchLoop)
		        	    datatable("Status", "dTestBatchExecution") = "FAILED"                             
		        	Case "NOT EXECUTED"
		        		datatable.GetSheet("dTestBatchExecution").SetCurrentRow(intTestBatchLoop)
		        	    datatable("Status", "dTestBatchExecution") = "NOT EXECUTED"
		        	Case "ABORTED"
		        		datatable.GetSheet("dTestBatchExecution").SetCurrentRow(intTestBatchLoop)
		        	    datatable("Status", "dTestBatchExecution") = "SKIPPED"			            
			    	End Select   				
				Exit For
			End If    		    	

	Elseif Ucase(strDependency) = Ucase("N") Then
		aDriver.RunTest(strTestCaseID)	
		datatable.GetSheet("dTestBatchExecution").setNextRow
	End If		
  Next
 'Next
Next
Set aDriver = Nothing
'msgbox (PathFinder.Locate ("..\..\Report\ResultLog.xls"))
DataTable.Export "..\..\Report\ResultLog.xls"
'DataTable.Export(Environment("TestDir") & "\" & "ResultLog" & ".xls")
'myobject1.Report_Final()
'datatable.Export "Resultlog.xls" @@ hightlight id_;_Browser("Google").Page("eBay").WebRadioGroup("LH BuyingFormats")_;_script infofile_;_ZIP::ssf4.xml_;_
'If strResultLogFullPath = "" Then
'	strResultLogFullPath = "..\Reports\Resultlog.xls"
'End If
'datatable.Export strResultLogFullPath

'Browser("Welcome to Maximo AWS").Page("Start Center").Link("Job Plans").Click

