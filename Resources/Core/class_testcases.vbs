Option Explicit
'*********************************************************
' TestCases CLASS
' Before making any change, please drop an email to arora.h@hcl.com
'*********************************************************
Class TestCases 

	'************************************************************************
	' Class Variables
	'************************************************************************
	Private mFileFullPath
    Private mSheetName
	Private mLocalSheetName
	Private mCurrentTest
	Private mFirstSubTest
	Private mSubTestCount

	'************************************************************************
	' Private Methods
	'************************************************************************
	
	' Methods for the class initialization

    Private Sub Class_Initialize() 
        mFileFullPath = ""
        mSheetName = ""
		mLocalSheetName = ""
	End Sub

	Private Sub Class_Terminate()
		' Nothing
	End Sub

    Public Sub Init (ByVal strFileFullPath, ByVal strSheetName) 
    
        mFileFullPath = strFileFullPath
        mSheetName = strSheetName
		mLocalSheetName = "d" + mSheetName
		datatable.AddSheet mLocalSheetName
		datatable.ImportSheet mFileFullPath, mSheetName, mLocalSheetName
	End Sub
	
	Private Function getFirstRow(ByVal strTestID)
		getFirstRow = -1
		Dim i, intNbRows  
		intNbRows = datatable.GetSheet(mLocalSheetName).GetRowCount
		For i=1 to intNbRows
			datatable.GetSheet(mLocalSheetName).SetCurrentRow(i)
			If datatable.Value("Test_ID",mLocalSheetName) = strTestID Then
				getFirstRow = i
				Exit For
			End If
		Next
	End Function

	Private Function getLastRow(ByVal strTestID)
		getLastRow = -1
		Dim i, j, intNbRows
		intNbRows = datatable.GetSheet(mLocalSheetName).GetRowCount
		'Search first step row concerning this test id
		For i=1 to intNbRows
			datatable.GetSheet(mLocalSheetName).SetCurrentRow(i)
			If datatable.Value("Test_ID",mLocalSheetName) = strTestID Then
				getLastRow = i
				Exit For
			End If
		Next
		'If one step has been found, loop until the last row for this test id
		If getLastRow <> -1 Then
			For j=i+1 to intNbRows
				datatable.GetSheet(mLocalSheetName).SetCurrentRow(j)
				If datatable.Value("Test_ID",mLocalSheetName) <> strTestID Then
					Exit For
				End If
				getLastRow = j
			Next
		End If
	End Function

	'************************************************************************
	' Public Methods
	'************************************************************************
	
	Public Function ExistTest(ByVal strTestID) 
        If getFirstRow(strTestID) = -1 Then
			ExistTest = False
		Else
			ExistTest = True
		End If
    End Function

	Public Function SetFirstTest()
        datatable.GetSheet(mLocalSheetName).SetCurrentRow(1)
    End Function

	Public Function GetTestCount()
		GetTestCount = datatable.GetSheet(mLocalSheetName).GetRowCount
	End Function

	Public Function SetCurrentTest(ByVal strTestID)
		mCurrentTest = strTestID
		mFirstSubTest = getFirstRow(strTestID)
		mSubTestCount = getLastRow(strTestID) - mFirstSubTest + 1
        datatable.GetSheet(mLocalSheetName).SetCurrentRow(mFirstSubTest)
    End Function
	
	Public Function SetFirstSubTest()
        datatable.GetSheet(mLocalSheetName).SetCurrentRow(mFirstSubTest)
    End Function

	Public Function SetNextSubTest()
		datatable.GetSheet(mLocalSheetName).SetNextRow
	End Function

	Public Function GetSubTestCount()
		GetSubTestCount = mSubTestCount
	End Function
    
	'*** Retrieve info of current Test and Sub Test
	Public Property Get ModuleID() 
        ModuleID = datatable.Value("Module_ID",mLocalSheetName)
'        environment("TestDataID")=datatable("query","dTestCases")
        'Changes for FWD	
'        msgbox environment("TestDataID")
    End Property
	
	Public Property Get TestID() 
        TestID = datatable.Value("Test_ID",mLocalSheetName)
        'If datatable.Value("Query1",mLocalSheetName) <> "" Then
        '	environment("TestDataID").value=environment("TestDataID").value & "_" &  datatable.Value("Query1",mLocalSheetName)	
        'End if
    End Property
	
	Public Property Get TestName() 
        TestName = datatable.Value("Test_Name",mLocalSheetName)
    End Property

	Public Property Get TestDescription() 
        TestDescription = datatable.Value("Test_Description",mLocalSheetName)
    End Property

	Public Property Get SubTestID() 
        SubTestID = datatable.Value("SubTest_ID",mLocalSheetName)
    End Property
	
	Public Property Get Environment() 
        Environment = datatable.Value("Environment",mLocalSheetName)
    End Property
	
	Public Property Get Data(ByVal strColumnID) 
        Data = datatable.Value(strColumnID,mLocalSheetName)
    End Property

	Public Property Get Execution() 
        Execution = datatable.Value("Execution",mLocalSheetName)
    End Property

	Public Property Get Status() 
        Status = datatable.Value("Status",mLocalSheetName)
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

	Public Property Let Status(ByVal strStatus) 
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

Public Function NewTestCases(ByVal strFileFullPath, ByVal strSheetName) 
'msgbox "strFileFullPath " & strFileFullPath
    Set NewTestCases = new TestCases 
    NewTestCases.Init strFileFullPath, strSheetName
End Function 
