Option Explicit
'*********************************************************
' WEB KEYWORDS
'*********************************************************
Class WebKeyword
	
	'*********************************************************
	' Purpose:  Run the action linked to a Web Keyword
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
	' Author: Hitesh Arora (arora.h@hcl.com)
	'*********************************************************
	Public Function runKeyword (ByVal strKeyword,ByVal strBrowserID,ByVal strPageID,ByVal strObjectID,ByVal strParam1,ByVal strParam2,ByVal strParam3) ' As Integer
		On Error Resume Next
		Err.Clear
	
		Dim arrReturn ' Array containing the result of the keyword function call
		Dim objCurrent
		Dim strObjectType
		Dim strCountry
	
		'Retrieve Object Type
		strObjectType = getObjectTypeFromKeyword(strKeyword)
		
		'Set strMainObj = Browser("title:=.*").Page("title:=.*")
	
		'Run Keyword
		Select Case strKeyword
		Case "web_browser_open"
			'Set objCurrent = getQTPObject(strObjectType, strBrowserID, strPageID, strObjectID)
			Set objCurrent = getQTPObject(strObjectType,"") 			
			arrReturn = web_browser_open(objCurrent, strParam1, strParam2)	
		Case "web_browser_close"
			Set objCurrent = getQTPObject(strObjectType,"") 
			arrReturn = web_browser_close(objCurrent)
		Case "web_page_exist"
			'Set objCurrent = getQTPObject(strObjectType, strBrowserID, strPageID, strObjectID)
			Set objCurrent = getQTPObject(strObjectType,"")			
			arrReturn = web_page_exist(objCurrent)
			
		Case "web_webelement_click"
			Set objCurrent = getQTPObject(strObjectType,strObjectID)
			arrReturn = web_webelement_click(objCurrent)
		
		Case "web_webelement_set"
			Set objCurrent = getQTPObject(strObjectType,strObjectID)
			arrReturn = web_webelement_set(objCurrent, strParam1)
			'arrReturn = generic_object_set(objCurrent, strParam1)
			
		Case "web_browser_wait"
			'Set objCurrent = getQTPObject(strObjectType, strBrowserID, strPageID, strObjectID) 
			arrReturn = generic_object_wait(strParam1)
			
		Case "web_link_click"
			'Set objCurrent = getQTPObject(strObjectType, strBrowserID, strPageID, strObjectID)
			Set objCurrent = getQTPObject(strObjectType, strMainObj,strObjectID)
			arrReturn = web_link_click(objCurrent)			
			'arrReturn = generic_object_click(objCurrent)	
			
		Case "web_page_country"
			Set objCurrent = getQTPObject(strObjectType, strMainObj,"") 
			arrReturn = web_page_country(objCurrent, strParam1)		


		Case "web_dialog_button_click"
			Set objCurrent = getQTPObject(strObjectType, strBrowserID, strPageID, strObjectID) 
			arrReturn = generic_object_exist(objCurrent, True)
			'arrReturn = generic_object_click(objCurrent)

		Case Else
			Err.Raise 1, "Keyword", strKeyword & " keyword not found"
			arrReturn = Array(1, "", Err.Description)
		End Select
		Set objCurrent = Nothing
	
		runKeyword = arrReturn
	End Function
	
	'*********************************************************
	' Object Repository search
	'*********************************************************
	
	Private Function getQTPObject(ByVal strObjectType,ByVal strObjectID) ' As Object
		Dim strMainObj,ObjDesc,stext
		Set strMainObj = Browser("title:=.*").Page("title:=.*")
		On Error Resume Next
		Err.Clear
		Set getQTPObject = Nothing
		If strObjectType = "browser" Then
			Set getQTPObject  = strMainObj
		ElseIf strObjectType = "page" Then
			Set getQTPObject  = strMainObj
		Else
			Set ObjDesc = description.Create	
			ObjDesc("xpath").value = strObjectID	
			If strMainObj.WebElement(ObjDesc).Exist(5) Then
			wait(3)
				stext = strMainObj.WebElement(ObjDesc).GetROProperty("micClass")
				If Ucase(stext) = Ucase("WebEdit") Then
					Set getQTPObject = strMainObj.WebEdit(ObjDesc)
				Else
					Set getQTPObject = strMainObj.WebElement(ObjDesc)	
				End If
			Else
				Set ObjDesc = description.Create
				ObjDesc("html id").value = strObjectID
				If strMainObj.WebElement(ObjDesc).Exist(2) Then
					stext = strMainObj.WebElement(ObjDesc).GetROProperty("micClass")
					If Ucase(stext) = Ucase("WebEdit") Then
						Set getQTPObject = strMainObj.WebEdit(ObjDesc)
					Else
						Set getQTPObject = strMainObj.WebElement(ObjDesc)	
					End If
				Else
					Set ObjDesc = description.Create
					ObjDesc("name").value = strObjectID
					If strMainObj.WebElement(ObjDesc).Exist(2) Then
						stext = strMainObj.WebElement(ObjDesc).GetROProperty("micClass")
						If Ucase(stext) = Ucase("WebEdit") Then
							Set getQTPObject = strMainObj.WebEdit(ObjDesc)
						Else
							Set getQTPObject = strMainObj.WebElement(ObjDesc)	
						End If
					End If
				End If
			End If			
		End If
		On Error GoTo 0
		If getQTPObject Is Nothing Then
			Err.Raise 1, "WebKeyword", strObjectID & " Object of type " & strObjectType & " not found"
		End If
	End Function
	
	'*********************************************************
	' Keyword implementations
	'*********************************************************

'****** Browser, Page, Dialog and Frame

	Private Function web_browser_open(ByRef objBrowser,ByVal strURL,ByVal strBrowser)
		'Open the Web Browser
		Select Case LCase(strBrowser)
			Case "ie"
                SystemUtil.Run "iexplore.exe", strURL,"","",9
			Case "firefox"
				'TODO: Test with Firefox Plugin 
				 SystemUtil.Run "firefox.exe", strURL,"","",7
			Case "chrome"
				'TODO: Test with Chrome Plugin 
				 SystemUtil.Run "chrome.exe", strURL,"","",7
			Case Else
				'If the Web Browser is not found
				web_browser_open = Array (1, "", "'" & strBrowser & "' is not a valide Web Browser.")
				Exit Function
		End Select

		'Check that the Web Browser is loaded
		objBrowser.Sync
		if objBrowser.Exist(4) = False Then
			Err.Raise 1, "Exist", "Browser not found"
			web_browser_open = Array (1, "", Err.Description)
		Else
			web_browser_open = Array (0, "", "Successfully completed.")
		End If
	End Function
	
	Private Function web_page_exist(ByRef objPage)
		objPage.Sync
		If objPage.Exist(4) = False Then
			Err.Raise 1, "Exist", "Page not found"
			web_
			page_exist = Array (1, "", Err.Description)
		Else
			web_page_exist = Array (0, "", "Successfully completed.")
		End If
	End Function
	
	Private Function web_webelement_click(ByRef objObject)
		objObject.Click
		web_webelement_click = Array (0, "", "Successfully completed.")
	End Function
	
	Private Function web_webelement_set(ByRef objObject, ByVal strValue)
		objObject.Set strValue
		web_webelement_set = Array (0, "", "Value '" & strValue & "' set in the field.")
	End Function
	
	Private Function web_browser_close(ByRef objBrowser)
		'If the Browser exists, close it
		if objBrowser.Exist(4) Then
			objBrowser.Close
			web_browser_close = Array (0, "", "Successfully completed.")
		Else
			web_browser_close = Array (1, "", "Browser not found.")
		End If
	End Function
	
	
	Private Function web_page_country(ByRef objObject,ByVal strCountry)
		Dim odesc,oChildObj,i,icnt,sText
		If objObject.Exist(3) Then
			Set odesc = description.Create
			odesc("micClass").value = "link"
			odesc("html tag").value = "A"
			Set oChildObj = objObject.ChildObjects(odesc)
			icnt = oChildObj.count
			'msgbox icnt
			For i = 0  To icnt
				wait(4)
				sText = oChildObj(i).GetROProperty("text")
				If instr(sText,strCountry) > 0 Then
					wait(4)
					oChildObj(i).Click
					wait(10)
					web_page_country = Array (0, "", "Successfully found "& strCountry & "link in the table.")
					Exit For
				End If	
			Next
		Else
			Err.Raise 1, "Exist", "Table does not contain" & strCountry & "link"
	        web_page_country = Array (1, "", Err.Description)
		End If
	End Function


	Private Function web_dialog_exist(ByRef objDialog)
		If objDialog.Exist(4) = False Then
			Err.Raise 1, "Exist", "Dialog not found"
			web_dialog_exist = Array (1, "", Err.Description)
		End If
		web_dialog_exist = Array (0, "", "Successfully completed.")
	End Function

	Private Function web_frame_exist(ByRef objFrame)
		If objFrame.Exist(4) = False Then
			Err.Raise 1, "Exist", "Frame not found"
			web_frame_exist = Array (1, "", Err.Description)
		End If
		web_frame_exist = Array (0, "", "Successfully completed.")
	End Function

'****** Image, Link, Static,ViewLink and WebButton

	Private Function web_image_click(ByRef objImage)
		objImage.Click
		web_image_click = Array (0, "", "Successfully completed.")
	End Function
	
	Private Function web_image_get(ByRef objImage, ByVal strValueID)
	   Dim strValue
	   strValue = objImage.GetROProperty("alt")
		Call saveRunValue (strValueID, strValue)
		 web_image_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function

	Private Function web_link_click(ByRef objLink)
		objLink.Click
		web_link_click = Array (0, "", "Successfully completed.")
	End Function
	

	Private Function web_link_get(ByRef objLink, ByVal strValueID)
	   Dim strValue
	   strValue = objLink.GetROProperty("innertext")
		Call saveRunValue (strValueID, strValue)
		 web_link_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function

	Private Function web_static_get(ByRef objStatic, ByVal strValueID)
		Dim strValue
		strValue =  objStatic.GetROProperty("text")
		Call saveRunValue (strValueID, strValue)
		web_webstatic_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function

	Private Function web_viewlink_click(ByRef objViewLink)
		objViewLink.Click
		web_viewlink_click = Array (0, "", "Successfully completed.")
	End Function
	
	Private Function web_viewlink_get(ByRef objViewLink, ByVal strValueID)
	   Dim strValue
	   strValue = objViewLink.GetROProperty("innertext")
		Call saveRunValue (strValueID, strValue)
		 web_viewlink_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function



	Private Function web_dialog_button_click(ByRef objButton)
		objButton.Click
		web_dialog_button_click = Array (0, "", "Successfully completed.")
	End Function
	
	Private Function web_webbutton_get(ByRef objButton, ByVal strValueID)
	   Dim strValue
	   strValue = objButton.GetROProperty("innertext")
		Call saveRunValue (strValueID, strValue)
		web_webbutton_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function

'********  WebCheckBox, WebEdit, WebElement, WebFile, WebList and WebRadioGroup

	Private Function web_webcheckbox_get(ByRef objCheckBox, ByVal strValueID)
	   Dim strValue
	   strValue = objCheckBox.GetROProperty("checked")
		Call saveRunValue (strValueID, strValue)
		 web_webcheckbox_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function
	Private Function web_webcheckbox_set(ByRef objCheckBox, ByVal strValue)
		objCheckBox.Set strValue
		web_webcheckbox_set = Array (0, "", "Value '" & strValue & "' selected in the CheckBox field.")
	End Function

	Private Function web_webedit_get(ByRef objEdit, ByVal strValueID)
	   Dim strValue
	   strValue = objEdit.GetROProperty("value")
		Call saveRunValue (strValueID, strValue)
		 web_webedit_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function
	
'	Private Function web_webelement_click(ByRef objWebElement)
'		objWebElement.Click
'		web_webelement_click = Array (0, "", "Successfully completed.")
'	End Function

	Private Function web_webelement_get(ByRef objWebElement, ByVal strValueID)
	   Dim strValue
	   strValue = objWebElement.GetROProperty("innertext")
		Call saveRunValue (strValueID, strValue)
		 web_webelement_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function

	Private Function web_weblist_get(ByRef objList, ByVal strValueID)
	   Dim strValue
	   strValue = objList.GetROProperty("value")
		Call saveRunValue (strValueID, strValue)
		 web_weblist_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function
	Private Function web_weblist_set(ByRef objList, ByVal strValue)
		objList.Select strValue
		web_weblist_set = Array (0, "", "Value '" & strValue & "' selected in the WebList field.")
	End Function

	Private Function web_webfile_get(ByRef objWebFile, ByVal strValueID)
	   Dim strValue
	   strValue = objWebFile.GetROProperty("value")
		Call saveRunValue (strValueID, strValue)
		 web_webfile_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function
	Private Function web_webfile_set(ByRef objWebFile, ByVal strValue)
		objWebFile.Select strValue
		web_webfile_set = Array (0, "", "Value '" & strValue & "' selected in the WebFile field.")
	End Function

	Private Function web_webradiogroup_get(ByRef objRadioGroup, ByVal strValueID)
	   Dim strValue
	   strValue = objRadioGroup.GetROProperty("value")
		Call saveRunValue (strValueID, strValue)
		 web_webradiogroup_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function
	Private Function web_webradiogroup_set(ByRef objRadioGroup, ByVal strValue)
		objRadioGroup.Select strValue
		web_webradiogroup_set = Array (0, "", "Value '" & strValue & "' selected in the RadioGroup field.")
	End Function

'********  WebTable

	Private Function web_webtable_get(ByRef objWebTable, ByVal strValueID)
	   Dim strValue
	   strValue = objWebTable.GetROProperty("innertext")
		Call saveRunValue (strValueID, strValue)
		 web_webtable_get = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function
	Private Function web_webtable_getcell(ByRef objWebTable, ByVal intRow, ByVal intColumn, ByVal strValueID)
	   Dim strValue
	   strValue = objWebTable.GetCellData(intRow,  intColumn)
		Call saveRunValue (strValueID, strValue)
		 web_webtable_getcell = Array (0, "", "Value '" & strValue & "' saved with the ValueID '" & strValueID & "'")
	End Function
	
End Class
