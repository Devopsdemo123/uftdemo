Option Explicit
'*********************************************************
' RUN KEYWORD FUNCTION
'*********************************************************
'*********************************************************
' Purpose:  Run the action linked to a Keyword
' Inputs:   objCurrent:       the data string to analyse (e.g. "Data(Query,Search)", "Env(URL)", "5")
'           strKeyword:       the keyword linked to the function to run
'           strBrowserID:    the id of the Browser in the ObjectRepository
'           strPageID:         the id of the Page in the ObjectRepository
'           strObjectID:       the id of the object in the ObjectRepository
'           strParam1:        first parameter to pass to the function to run (optional)
'           strParam2:        second parameter to pass to the function to run (optional)
'           strParam3:        third parameter to pass to the function to run (optional)
' Returns:  The return code of the keyword function. 
'           If the keyword hasn't been found, returns 1 and raise an error.
' Author: Hitesh Arora (HCL Tech)
'*********************************************************
Public Function runKeyword (ByVal strKeyword,ByVal strBrowserID,ByVal strPageID,ByVal strObjectID,ByVal strParam1,ByVal strParam2,ByVal strParam3) ' As Integer
	On Error Resume Next
	Err.Clear

	Dim arrReturn

	' Call the Keyword run function of the concerned library
	Dim strLibrary
	strLibrary = getLibraryFromKeyword(strKeyword)
	If Err.Number <> 0 Then
		arrReturn = Array (1, Err.Description)
		Exit Function
	End If

	Dim objKeywordLibrary

	Select Case strLibrary
	Case "common"
		Set objKeywordLibrary = New CommonKeyword
	Case "web"
		Set objKeywordLibrary = New WebKeyword
	Case "java"
		Set objKeywordLibrary = New JavaKeyword
	Case "oracle"
		Set objKeywordLibrary = New OracleKeyword
	Case "DBValidation"
		Set objKeywordLibrary = New DataBase
	'############# Add here other Keyword Libraries ###########
'	Case "<prefix>"
'		Set objKeywordLibrary = New <Prefix>Keyword
	'#####################################################
	Case Else
		'If Library not found
		Err.Raise 1, "Keyword",  "'" & strLibrary & "' keyword library for keyword '" & strKeyword &"' not found"
		arrReturn = Array (1, Err.Description)
		runKeyword = arrReturn
		Exit Function
	End Select

	'Run the Keyword 
	arrReturn = objKeywordLibrary.runKeyword (strKeyword, strBrowserID, strPageID, strObjectID, strParam1, strParam2, strParam3)

	Set objKeywordLibrary = Nothing

	runKeyword = arrReturn
End Function

'*********************************************************
' KEYWORD UTILITY FUNCTIONS
'*********************************************************

'*********************************************************
' Purpose:  Retrieves Object type from the keyword. A Keyword
' Inputs:   strKeyword:      the keyword following the format <library>_<object type>_<action>
' Returns:  The object type or return the found in the Object Repository. 
'           If the object isn't found, returns Nothing and raise an error.
'*********************************************************
Public Function getObjectTypeFromKeyword(ByVal strKeyword) ' As String
	On Error Resume Next
	Err.Clear
	getObjectTypeFromKeyword = ""
	
	Dim arrSplit
	arrSplit = Split (strKeyword, "_")
	If Ubound(arrSplit) > 1 Then
		getObjectTypeFromKeyword = arrSplit(1)
	Else
		Err.Raise 1, "Keyword", "'" & strKeyword & "' keyword doesn't contain an object type"
	End If
End Function

'*********************************************************
' Purpose:  Retrieves Library from the keyword. A Keyword
' Inputs:   strKeyword:      the keyword following the format <library>_><object type>_<action>
' Returns:  The object type or return the found in the Object Repository. 
'           If the object isn't found, returns Nothing and raise an error.
'*********************************************************
Public Function getLibraryFromKeyword(ByVal strKeyword) ' As String
	On Error Resume Next
	Err.Clear
	getLibraryFromKeyword = ""
	
	Dim arrSplit
	arrSplit = Split (strKeyword, "_")
	If Ubound(arrSplit) > 1 Then
		getLibraryFromKeyword = arrSplit(0)
	Else
		Err.Raise 1, "Keyword", "'" & strKeyword & "' keyword doesn't contain a library name"
	End If
End Function

Public Sub saveRunValue(ByVal strValueID, ByVal strValue)
	Environment.Value("temp_" & strValueID) = strValue
End Sub

'*********************************************************
' Generic KEYWORD Implementation
'*********************************************************
Public Function generic_object_exist(ByRef objObject, ByRef blnAbordTest)
	If objObject.Exist(4) = False Then
		Dim strErrorDescription
		strErrorDescription = objObject.GetROPRoperty("micClass") & " not found"
		If blnAbordTest Then
			Err.Raise 1, "Error", strErrorDescription
		End If
		generic_object_exist = Array (1, "", strErrorDescription)
	Else
        generic_object_exist = Array (0, "", "Successfully completed.")
	End If
End Function
Public Function generic_object_activate(ByRef objObject, ByRef blnAbordTest)
	If objObject.Exist(4) = False Then
		Dim strErrorDescription
		strErrorDescription = objObject.GetROPRoperty("micClass") & " not found"
		If blnAbordTest Then
			Err.Raise 1, "Error", strErrorDescription
		End If
		generic_object_activate = Array (1, "", strErrorDescription)
	Else
		objObject.Activate
        generic_object_activate = Array (0, "", "Successfully completed.")
	End If
End Function
Public Function generic_object_wait(ByRef intValueID)
	wait(intValueID)
	generic_object_wait = Array (0, "", "Successfully paused for " & intValueID)
End Function
Public Function generic_object_click(ByRef objObject)
	objObject.Click
	generic_object_click = Array (0, "", "Successfully click on the " & objObject.GetROPRoperty("micClass"))
End Function
Public Function generic_object_get(ByRef objObject, ByVal strValueID, ByVal strPropertyID)
	Dim strValue
	strValue = objObject.GetROProperty(strPropertyID)
	Call saveRunValue (strValueID, strValue)
	generic_object_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
End Function
Public Function generic_object_select(ByRef objObject, ByVal strValue)
	objObject.Select strValue
	generic_object_select = Array (0, "", "Value '" & strValue & "' selected in the " & objObject.GetROPRoperty("micClass") & " field")
End Function
Public Function generic_object_set(ByRef objObject, ByVal strValue)
	objObject.Set strValue
	generic_object_set = Array (0, "", "Value '" & strValue & "' selected in the " & objObject.GetROPRoperty("micClass") & " field")
End Function
Public Function generic_object_sendkey_enter(ByRef objObject)
   objObject.Click
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.SendKeys "{ENTER}"
	wait(5)
	Set objShell = Nothing
	generic_object_sendkey_enter = Array (0, "", "The 'Enter' key has been pressed on the " & objObject.GetROPRoperty("micClass") & " field")
End Function

'*********************************************************
' Common KEYWORDS
'*********************************************************

Class CommonKeyword
	
	'*********************************************************
	' Purpose:  Run the action linked to a Common Keyword
	' Inputs:   
	'           strKeyword:       the keyword linked to the function to run
	'           strBrowserID:    the id of the Browser in the ObjectRepository
	'           strPageID:         the id of the Page in the ObjectRepository
	'           strObjectID:       the id of the object in the ObjectRepository
	'           strParam1:        first parameter to pass to the function to run (optional)
	'           strParam2:        second parameter to pass to the function to run (optional)
	'           strParam3:        third parameter to pass to the function to run (optional)
	' Returns:  The return code of the keyword function. 
	'           If the keyword hasn't been found, returns 1 and raise an error.
	'*********************************************************
	Public Function runKeyword (ByVal strKeyword,ByVal strBrowserID,ByVal strPageID,ByVal strObjectID,ByVal strParam1,ByVal strParam2,ByVal strParam3) ' As Integer
		On Error Resume Next
		Err.Clear
	
		Dim arrReturn ' Array containing the result of the keyword function call
	
		Dim objCurrent
		Dim strObjectType
	
		'Retrieve Object Type
		strObjectType = getObjectTypeFromKeyword(strKeyword)
	
		'Run Keyword
		Select Case strKeyword
		Case "common_check_equal"
			arrReturn = common_check_equal(strParam1, strParam2)
		Case "common_check_equal_trim"
			arrReturn = common_check_equal_trim(strParam1, strParam2)
		Case "common_check_notequal"
			arrReturn = common_check_notequal(strParam1, strParam2)
		Case "common_check_notequal_trim"
			arrReturn = common_check_notequal_trim(strParam1, strParam2)
		Case "common_check_regexp"
			arrReturn = common_check_regexp(strParam1, strParam2)
		Case "common_check_notequal_regexp"
			arrReturn = common_check_notequal_regexp(strParam1, strParam2)
		Case Else
			Err.Raise 1, "Common Keyword", strKeyword & " keyword not found"
			arrReturn = Array(1, "", Err.Description)
		End Select
		Set objCurrent = Nothing
	
		runKeyword = arrReturn
	End Function
	
	'*********************************************************
	' Object Repository search
	'*********************************************************
	
	Private Function getQTPObject(ByVal strObjectType,ByVal strBrowserID,ByVal strPageID,ByVal strObjectID) ' As Object
		On Error Resume Next
		Err.Clear
		Set getQTPObject = Nothing
		Select Case strObjectType
	'	Case "browser"
	'		Set getQTPObject  = Browser(strObjectID)
		End Select
		On Error GoTo 0
		If getQTPObject Is Nothing Then
			Err.Raise 1, "Obkject Repository", strObjectID & " Object of type " & strObjectType & " not found in the Object Repository"
		End If
	End Function
	
	'*********************************************************
	' Keyword implementations
	'*********************************************************

	Private Function common_check_equal(ByVal varExpectedValue, ByVal varActualValue)
		If varExpectedValue = varActualValue Then
			common_check_equal = Array (0, "Equal '" & varExpectedValue & "'", varActualValue)
		Else
			common_check_equal = Array (1, "Equal '" & varExpectedValue & "'", varActualValue)
		End If
	End Function
	Private Function common_check_equal_trim(ByVal varExpectedValue, ByVal varActualValue)
		If varExpectedValue = Trim(varActualValue) Then
			common_check_equal_trim = Array (0, "Equal '" & varExpectedValue & "'", varActualValue)
		Else
			common_check_equal_trim = Array (1, "Equal '" & varExpectedValue & "'", varActualValue)
		End If
	End Function
	Private Function common_check_notequal(ByVal varExpectedValue, ByVal varActualValue)
		If varExpectedValue <> varActualValue Then
			common_check_notequal = Array (0, "Not equal '" & varExpectedValue & "'", varActualValue)
		Else
			common_check_notequal = Array (1, "Not equal '" & varExpectedValue & "'", varActualValue)
		End If
	End Function
	Private Function common_check_notequal_trim(ByVal varExpectedValue, ByVal varActualValue)
		If varExpectedValue <> Trim(varActualValue) Then
			common_check_notequal_trim = Array (0, "Not equal '" & varExpectedValue & "'", varActualValue)
		Else
			common_check_notequal_trim = Array (1, "Not equal '" & varExpectedValue & "'", varActualValue)
		End If
	End Function
	Private Function common_check_regexp(ByVal varExpectedPattern, ByVal varActualValue)
		Dim objRegExp
		Set objRegExp = New RegExp
		objRegExp.Pattern = varExpectedPattern
		objRegExp.IgnoreCase = False ' Set case sensitivity.
		If objRegExp.Test(varActualValue) Then
			common_check_regexp = Array (0, "Match the pattern '" & varExpectedPattern & "'", "The Value '" & varActualValue & "' matches the Pattern")
		Else
			common_check_regexp = Array (1, "Match the pattern '" & varExpectedPattern & "'","The Value '" & varActualValue & "' doesn't match the Pattern")
		End If
		Set objRegExp = Nothing
	End Function
Private Function common_check_notequal_regexp(ByVal varExpectedPattern, ByVal varActualValue)
		Dim objRegExp
		Set objRegExp = New RegExp
		objRegExp.Pattern = varExpectedPattern
		objRegExp.IgnoreCase = False ' Set case sensitivity.
		If not objRegExp.Test(varActualValue) Then
			common_check_regexp = Array (0, "Negative Test Match the pattern '" & varExpectedPattern & "'", "The Value '" & varActualValue & "' does not matches the Pattern")
		Else
			common_check_regexp = Array (1, "Negative Test Match the pattern '" & varExpectedPattern & "'","The Value '" & varActualValue & "' match the Pattern")
		End If
		Set objRegExp = Nothing
	End Function
End Class
