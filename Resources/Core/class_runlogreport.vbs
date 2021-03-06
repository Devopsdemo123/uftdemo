Option Explicit
'*********************************************************
' RunLogReport CLASS
' Before making any change, please drop an email to arora.h@hcl.com
'*********************************************************
Class RunLogReport

	'*********************************************************
	' ATTRIBUTS
	'*********************************************************
	Private mFileFullPath
	Private mSheetName
	Private mLocalSheetName

	Private mCurrentTestId
	Private	mCurrentTestName
	Private mCurrentTestDesc
	
	'*********************************************************
	' Initialize/Terminate METHODS
	'*********************************************************
	
    Private Sub Class_Initialize() 
		mFileFullPath = ""
        mSheetName = ""
		mLocalSheetName = ""

		mCurrentTestId = ""
		mCurrentTestName = ""
		mCurrentTestDesc = ""
	End Sub

	Private Sub Class_Terminate()
		' Nothing
	End Sub

	 Public Sub Init(ByVal strFileFullPath,ByVal strSheetName) 
		mFileFullPath = strFileFullPath
        mSheetName = strSheetName
		mLocalSheetName = "d" + mSheetName
		datatable.AddSheet mLocalSheetName
		datatable.ImportSheet mFileFullPath, mSheetName, mLocalSheetName
	End Sub

	'#################

	'*********************************************************
	' PUBLIC METHODS
	'*********************************************************
	Public Sub TestBegin(ByVal strTestId, ByVal strTestName, ByVal strTestDesc)
	   'Store the values for Columns: Test_ID, Test_Name, Test_Description

	   mCurrentTestId = strTestId
	   mCurrentTestName = strTestName
	   mCurrentTestDesc = strTestDesc
	End Sub

	Public Sub StepBegin(ByVal strSubTestID, ByVal strSubTestDesc, ByVal strStepID, ByVal strStepDesc, _
                                                    ByVal strKeyword, ByVal strBrowserID, ByVal strPageID, ByVal strObjectID, _
                                                    ByVal strParam1, ByVal strParam2, ByVal strParam3, _
                                                    ByVal strExecution, ByVal strStatus, ByVal strStartDateTime)
		'Set the values of the columns: Test_ID, Test_Name, Test_Description, 
		'SubTest_ID	SubTest_Description	Step_ID	Step_Description	Keyword	Browser_ID	Page_ID	Object_ID	Param1	Param2	Param3
		'Execution	Status	Start_DateTime
		datatable.Value("Test_ID",mLocalSheetName) = mCurrentTestId
		datatable.Value("Test_Name",mLocalSheetName) = mCurrentTestName
		datatable.Value("Test_Description",mLocalSheetName) = mCurrentTestDesc
		datatable.Value("SubTest_ID",mLocalSheetName) = strSubTestID
		datatable.Value("SubTest_Description",mLocalSheetName) = strSubTestDesc
		datatable.Value("Step_ID",mLocalSheetName) = strStepID
		datatable.Value("Step_Description",mLocalSheetName) = strStepDesc
		datatable.Value("Keyword",mLocalSheetName) = strKeyword
		datatable.Value("Browser_ID",mLocalSheetName) = strBrowserID
		datatable.Value("Page_ID",mLocalSheetName) = strPageID
		datatable.Value("Object_ID",mLocalSheetName) = strObjectID
		datatable.Value("Param1",mLocalSheetName) = strParam1
		datatable.Value("Param2",mLocalSheetName) = strParam2
		datatable.Value("Param3",mLocalSheetName) = strParam3
		datatable.Value("Execution",mLocalSheetName) = strExecution
		datatable.Value("Status",mLocalSheetName) = strStatus
		datatable.Value("Start_DateTime",mLocalSheetName) = strStartDateTime 
	End Sub

	Public Sub UpdateKeywordParameters(ByVal strParam1, ByVal strParam2, ByVal strParam3)
		'Set the values of the columns: Param1	Param2	Param3
		datatable.Value("Param1",mLocalSheetName) = strParam1
		datatable.Value("Param2",mLocalSheetName) = strParam2
		datatable.Value("Param3",mLocalSheetName) = strParam3
	End Sub

	Public Sub StepEnd(ByVal strStatus,ByVal strEndDateTime, ByVal strExpectedResult,ByVal strActualResult, ByVal strError)
		' Set the values of the columns: Status End_DateTime	Expected_Result	Actual_Result	Error
        datatable.Value("Status",mLocalSheetName) = strStatus
		datatable.Value("End_DateTime",mLocalSheetName) = strEndDateTime
		datatable.Value("Expected_Result",mLocalSheetName) = strExpectedResult
		datatable.Value("Actual_Result",mLocalSheetName) = strActualResult
		datatable.Value("Error",mLocalSheetName) = strError

		'Move to next Row
		datatable.GetSheet(mLocalSheetName).SetNextRow
	End Sub

	Public Sub TestEnd()
		'Nothing
	End Sub

End Class

'*********************************************************
' CONSTRUCTION FUNCTION
'*********************************************************

Public Function NewRunLogReport(ByVal strFileFullPath,ByVal strSheetName) 
    Set NewRunLogReport = new RunLogReport
	NewRunLogReport.Init strFileFullPath, strSheetName
End Function
