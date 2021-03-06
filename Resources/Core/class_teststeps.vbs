Option Explicit
'*********************************************************
' TestSteps CLASS
' Before making any change, please drop an email to arora.h@hcl.com
'*********************************************************
Class TestSteps

	'************************************************************************
	' Class Variables
	'************************************************************************
	Private mFileFullPath
	Private mSheetName
	Private mLocalSheetName
	Private mCurrentSubTest
	Private mFirstStep
	Private mStepCount

	'************************************************************************
	' Private Methods
	'************************************************************************
	
	' Methods for the class initialization

    Private Sub Class_Initialize() 
		mFileFullPath = ""
        mSheetName = ""
		mLocalSheetName = ""
		mCurrentSubTest = ""
		mFirstStep = ""
        mStepCount = ""
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

	' Other private methods

	'**********************************************************************************************************************
	'Function/Subroutine Name : getFirstRow
	'Description : This function retrieves the row number of the first step concerning the test id
	'Inputs : strSubTestID: Id of the test
	'Output : Integer containing the row number of the first step concerning the test or -1 if test id not found

	'***********************************************************************************************************************
	Private Function getFirstRow(ByVal strSubTestID)
		getFirstRow = -1
		Dim i, intNbRows  
		intNbRows = datatable.GetSheet(mLocalSheetName).GetRowCount
		For i=1 to intNbRows
			datatable.GetSheet(mLocalSheetName).SetCurrentRow(i)
			If datatable.Value("SubTest_ID",mLocalSheetName) = strSubTestID Then
				getFirstRow = i
				Exit For
			End If
		Next
	End Function

	'**********************************************************************************************************************
	'Function/Subroutine Name : getLastRow
	'Description : This function retrieves the row number of the last step concerning the test id
	'Inputs : strSubTestID: Id of the test
	'Output : Integer containing the row number of the last step concerning the test or -1 if test id not found

	'***********************************************************************************************************************
	Private Function getLastRow(ByVal strSubTestID)
		getLastRow = -1
		Dim i, j, intNbRows  
		intNbRows = datatable.GetSheet(mLocalSheetName).GetRowCount
		'Search first step row concerning this test id
		For i=1 to intNbRows
			datatable.GetSheet(mLocalSheetName).SetCurrentRow(i)
			If datatable.Value("SubTest_ID",mLocalSheetName) = strSubTestID Then
				getLastRow = i
				Exit For
			End If
		Next
		'If one step has been found, loop until the last row for this test id
		If getLastRow <> -1 Then
			For j=i+1 to intNbRows
				datatable.GetSheet(mLocalSheetName).SetCurrentRow(j)
				If datatable.Value("SubTest_ID",mLocalSheetName) <> strSubTestID Then
					Exit For
				End If
				getLastRow = j
			Next
		End If
	End Function

	'************************************************************************
	' Public Methods
	'************************************************************************

	'**********************************************************************************************************************
	'Function/Subroutine Name : ExistTest
	'Description : This function checks whether the test id exists in the steps table
	'Inputs : strSubTestID: Id of the test
	'Output : Boolean containing true if the test exists and false if it doesn't.

	'***********************************************************************************************************************
	Public Function ExistTest(ByVal strSubTestID) 
        If getFirstRow(strSubTestID) = -1 Then
			ExistTest = False
		Else
			ExistTest = True
		End If
    End Function

	'**********************************************************************************************************************
	'Function/Subroutine Name : SetCurrentTest
	'Description : This function assigned the current test and put the cursor on the first step of the test
	'Prerequesite : Call the function ExistTest(strSubTestID) before.
	'Inputs : strSubTestID: Id of the test
	'Output : Nil

	'***********************************************************************************************************************
    Public Function SetCurrentSubTest(ByVal strSubTestID)
		mCurrentSubTest = strSubTestID
		mFirstStep = getFirstRow(strSubTestID)
		mStepCount = getLastRow(strSubTestID) - mFirstStep + 1
        datatable.GetSheet(mLocalSheetName).SetCurrentRow(mFirstStep)
    End Function
	
	'**********************************************************************************************************************
	'Function/Subroutine Name : SetCurrentTest
	'Description : Put the cursor on the first step of the test
	'Prerequesite : Call the function SetFirstStep(strSubTestID) before.
	'Inputs : strSubTestID: Id of the test
	'Output : Nil

	'***********************************************************************************************************************
	Public Function SetFirstStep()
        datatable.GetSheet(mLocalSheetName).SetCurrentRow(mFirstStep)
    End Function

	Public Function SetNextStep()
		datatable.GetSheet(mLocalSheetName).SetNextRow
	End Function

	Public Function GetStepCount()
		GetStepCount = mStepCount
	End Function
    
	'*** retrieve info of current test step
	Public Property Get SubTestID() 
        SubTestID = datatable.Value("SubTest_ID",mLocalSheetName)
    End Property

	Public Property Get SubTestDescription() 
        SubTestDescription = datatable.Value("SubTest_Description",mLocalSheetName)
    End Property
	
	Public Property Get StepID() 
        StepID = datatable.Value("Step_ID",mLocalSheetName)
    End Property
	
	Public Property Get StepDescription() 
        StepDescription = datatable.Value("Step_Description",mLocalSheetName)
    End Property

	Public Property Get Keyword() 
        Keyword = datatable.Value("Keyword",mLocalSheetName)
    End Property

	Public Property Get BrowserID() 
        BrowserID = datatable.Value("Browser_ID",mLocalSheetName)
    End Property

	Public Property Get PageID() 
        PageID = datatable.Value("Page_ID",mLocalSheetName)
    End Property
    	
	Public Property Get ObjectID() 
        ObjectID = datatable.Value("Object_ID",mLocalSheetName)
    End Property
	
	Public Property Get Param(ByVal strParamID) 
        Param = datatable.Value("Param" & strParamID,mLocalSheetName)
    End Property

	Public Property Get Screenshot() 
        Screenshot = datatable.Value("Screenshot",mLocalSheetName)
    End Property

	Public Property Get Execution() 
        Execution = datatable.Value("Execution",mLocalSheetName)
    End Property
	
	Public Property Get Status() 
        Status = datatable.Value("Status",mLocalSheetName)
    End Property
	
	Public Property Get Return() 
        Return = datatable.Value("Return",mLocalSheetName)
    End Property

	Public Property Get Error() 
        Error = datatable.Value("Error",mLocalSheetName)
    End Property

	Public Property Get StartDateTime() 
        StartDateTime = datatable.Value("Start_DateTime",mLocalSheetName)
    End Property

	Public Property Get EndDateTime() 
        EndDateTime = datatable.Value("End_DateTime",mLocalSheetName)
    End Property

	Public Property Let Status(strStatus) 
        datatable.Value("Status",mLocalSheetName) = strStatus
    End Property
	
	Public Property Let Return(ByVal strReturn) 
        datatable.Value("Return",mLocalSheetName) = strReturn
    End Property

	Public Property Let Error(ByVal strError) 
        datatable.Value("Error",mLocalSheetName) = strError
    End Property

	Public Property Let StartDateTime(ByVal strStartDateTime) 
        datatable.Value("Start_DateTime",mLocalSheetName) = strStartDateTime
    End Property

	Public Property Let EndDateTime(ByVal strEndDateTime) 
        datatable.Value("End_DateTime",mLocalSheetName) = strEndDateTime
    End Property
	
End Class

Public Function NewTestSteps(ByVal strFileFullPath,ByVal strSheetName) 
    Set NewTestSteps = new TestSteps 
	NewTestSteps.Init strFileFullPath, strSheetName
End Function
