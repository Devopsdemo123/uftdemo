Option Explicit
'*********************************************************
' TestData CLASS
' Before making any change, please drop an email to arora.h@hcl.com
'*********************************************************
Class TestData 

	'************************************************************************
	' Class Variables
	'************************************************************************
    Private mFileFullPath
    Private mSheetName
	Private mLocalSheetName
	Private mCurrentTest

	'************************************************************************
	' Private Methods
	'************************************************************************
	
	' Methods for the class initialization

    Private Sub Class_Initialize() 
        mFileFullPath = ""
        mSheetName = ""
		mLocalSheetName = ""
    End Sub

	Public Sub Init(ByVal strFileFullPath, ByVal strSheetName) 
        mFileFullPath = strFileFullPath
        mSheetName = strSheetName
		mLocalSheetName = "d" + mSheetName
		If existDataSheet(mLocalSheetName) = False Then
			datatable.AddSheet mLocalSheetName
			datatable.ImportSheet mFileFullPath, mSheetName, mLocalSheetName
		End If
    End Sub

	Private Sub Class_Terminate()
		' Nothing
	End Sub

	' Utility Functions
    Private Function getFirstRow(ByVal strDataID)
		getFirstRow = -1
		Dim i, intNbRows  
		intNbRows = datatable.GetSheet(mLocalSheetName).GetRowCount
		For i=1 to intNbRows
			datatable.GetSheet(mLocalSheetName).SetCurrentRow(i)
			If datatable.Value("Data_ID",mLocalSheetName) = strDataID Then
				getFirstRow = i
				Exit For
			End If
		Next
	End Function

    'Function to check if a DataTable sheet exists or not
	Private Function existDataSheet(ByVal strSheetName)
	    existDataSheet = True
		On Error Resume Next
		Dim objTest
		Set objTest = DataTable.GetSheet(strSheetName)
		If Err.Number Then
			existDataSheet = False
		End If
        On Error Goto 0
	End Function

	'************************************************************************
	' Public Methods
	'************************************************************************
	
	Public Function ExistData(ByVal strDataID) 
        If getFirstRow(strDataID) = -1 Then
			ExistData = False
		Else
			ExistData = True
		End If
    End Function

    Public Sub SetCurrentData(ByVal strDataID)
	   Dim intDataRow
		intDataRow = getFirstRow(strDataID)	
        datatable.GetSheet(mLocalSheetName).SetCurrentRow(intDataRow)
    End Sub
	 
	'*** retrieve info of current test
	Public Property Get DataID() 
        DataID = datatable.Value("Data_ID",mLocalSheetName)
    End Property
	
	Public Property Get DataDescription() 
        DataDescription = datatable.Value("Data_Description",mLocalSheetName)
    End Property
	
	Public Property Get Data(strColumnID) 
        Data = datatable.Value(strColumnID,mLocalSheetName)
    End Property
	
End Class 

'************************************************************************
' Function to construct the object
'************************************************************************
Public Function NewTestData(ByVal strFileFullPath, ByVal strSheetName) 
    Set NewTestData = new TestData 
    NewTestData.Init strFileFullPath, strSheetName
End Function