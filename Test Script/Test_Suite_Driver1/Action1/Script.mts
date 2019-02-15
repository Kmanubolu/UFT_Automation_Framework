'===========================t==SearchRelatedAccount
'=================================================
'Purpose: Main Driver Script
'==============================================================================
'Environment.LoadFromFile Parameter("EnvironmentFilePath")
'LoadFunctionLibrary Environment.Value("FrameworkFolderPath") & "\Function Library\IOFunctions.vbs"
Call fnLoadLibraries() 




'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'Set objQuoteAlert = objParent.WebElement("class:=x-container x-container-default x-table-layout-ct","innertext:=This quote will require underwriting approval prior to binding\.")
'objQuoteAlert.Highlight
'Set objRiskAnalysisTab = objParent.WebElement("class:=x-tree-node-text ","innertext:=Risk Analysis")
'objRiskAnalysisTab.Highlight
''fnTableSelectSpecifiedCheckBox(objTab,strValue)
''fnSelectTableRadioBtn(objTab,strValue)
'
'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'Set objRiskAnalysisIssuesTable = objParent.WebTable("class:=x-gridview-.*-table x-grid-table x-grid-with-col-lines x-grid-with-row-lines","cols:=5")
'objRiskAnalysisIssuesTable.Highlight
'rowCnt = objRiskAnalysisIssuesTable.RowCount
'For Iterator = 2 To rowCnt Step 1
'   Setting.WebPackage("ReplayType") = 2	
'	objRiskAnalysisIssuesTable.ChildItem(Iterator,1,"Image",0).Click
'	Setting.WebPackage("ReplayType") = 1
'	Wait 1
'Next
'
'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'Set objRiskAnalysisApprove = objParent.WebElement("class:=x-btn-inner x-btn-inner-center", "innertext:=Approve")
'objRiskAnalysisApprove.Highlight
'
'Set objParent = Browser("name:=.*Guidewire PolicyCenter.*").Page("title:=.*Guidewire PolicyCenter.*")
'Set objRiskAnalysisOK= objParent.WebElement("class:=x-btn-inner x-btn-inner-center","innertext:=OK")
''Set objRiskAnalysisOK= objParent.WebElement("html id:=RiskApprovalDetailsPopup:Update-btnInnerEl")
'objRiskAnalysisOK.Highlight
'
'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'Set objTreeQuote = objParent.WebElement("class:=x-tree-node-text ","innertext:=Quote")
'objTreeQuote.Highlight
'






'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'Set objTable = objParent.WebTable("class:=x-gridview-1300-table x-grid-table x-grid-with-col-lines x-grid-with-row-lines","cols:=5")
'blnFlag = fnClickSubmissionNavigate(objTable,"30655882")
'msgbox blnFlag
'Public Function fnClickSubmissionNavigate(objTable,strSelectionValue)

'	blnFlag = False
'	intRows = objTable.RowCount
'	If intRows > 0 Then
'	   For intI = 1 to intRows     
'		  If Trim(objTable.GetCellData(intI,3)) = Trim(strSelectionValue) Then
'		  	 Environment("SubmissionStatus") = objTable.GetCellData(intI,6)
'			 Setting.WebPackage("ReplayType") = 2
'			 objTable.ChildItem(intI,3,"WebElement",1).Click
'			 Setting.WebPackage("ReplayType") = 1		 
'			 Exit For		
'		  End If
'	   Next
'	End If
'	
'	Select Case Environment("SubmissionStatus")
'		Case "Draft","Not-taken","Withdrawn","Expired"
'		   Set objToVerify = objEleQualification
'		Case "Quoted","Bound"
'		   Set objToVerify = objEleQuoteTitle
'	End Select
'	
'	'Verify whether page navigation to expected page is successful
'	If objToVerify.Exist(Environment("SyncTimeOut")) Then  
'	 	blnFlag = True
'	End If
'	
'	fnClickSubmissionNavigate = blnFlag
'End Function

'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'Set objTable = objParent.WebTable("class:=.*x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","name:=infobar_WC","index:=0")


'blnFlag = fnTableClickSpecifiedItemNavigate("TransactionType","Submission")
'msgbox blnFlag
'
'Public Function fnTableClickSpecifiedItemNavigate(strSelectionType,strSelectionValue)
'	blnFlag = False	
'	Select Case strSelectionType
'		Case "Policy"
'			indx = 0
'			intCol = 1
'			intStatusCol = 3
'		Case "Transaction" 
'			indx = 1
'			intCol = 2
'			intStatusCol = 3
'		Case "TransactionType"	
'			indx = 1
'			intCol = 6
'			intStatusCol = 3
'	End Select
'	Set objParent = Browser("name:=.*Guidewire PolicyCenter.*").Page("title:=.*Guidewire PolicyCenter.*")
'	Set objTable = objParent.WebTable("class:=.*x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","name:=WebTable","index:="&indx)
'	intRows = objTable.RowCount
'	If intRows > 0 Then
'	   For intI = 1 to intRows     
'		  If Trim(objTable.GetCellData(intI,intCol)) = Trim(strSelectionValue) Then
'		  	 Environment("PolicyTransTypeStatus") = objTable.GetCellData(intI,intStatusCol)
'			 Setting.WebPackage("ReplayType") = 2
'			 objTable.ChildItem(intI,intCol,"WebElement",1).Click
'			 Setting.WebPackage("ReplayType") = 1		 
'			 Exit For		
'		  End If
'	   Next
'	End If
'	
'	Select Case Environment("PolicyTransTypeStatus")
'		Case "Draft"
'		   Set objToVerify = objEleQualification
'		Case "Quoted"
'		   Set objToVerify = objEleQuoteTitle	
'		Case "Canceled","In Force"
'		   Set objToVerify = objEleSummary
'	End Select
'	
'	'Verify whether page navigation to expected page is successful
'	If objToVerify.Exist(Environment("SyncTimeOut")) Then  
'	 	blnFlag = True
'	End If
'	
'	fnTableClickSpecifiedItemNavigate = blnFlag
'End Function
'





''objParent.WebTable("class:=.*x-grid-table x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","index:=3").Highlight
'
'intRows = objTable.RowCount
'If intRows > 0 Then
'	For intI = 1 to intRows
'		If Trim(objTable.GetCellData(intI,1)) = Trim(strPolicyNo) Then
'			Setting.WebPackage("ReplayType") = 2
'			objTable.ChildItem(intI,1,"WebElement",1).Click
'			Setting.WebPackage("ReplayType") = 1
'			Exit For
'		End If
'	Next	
'End If
'

'objParent.WebElement("html id:=.*AccountFile_LocationsLV_tb:setToPrimary-btnInnerEl","innertext:=Set As Primary").Highlight
'objParent.WebTable("html id:=gridview.*","class:=.*x-grid-table x-grid-with-col-lines x-grid-with-row-lines.*","index:=0").Highlight
'objParent.WebElement("class:=x-tree-node-text","innertext:=Locations","index:=0").Highlight
'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'Set objTable = objParent.WebTable("html id:=gridview.*","class:=.*x-grid-table x-grid-with-col-lines x-grid-with-row-lines.*","index:=0")
'strStatus = fnVerifyActiveStatus(objTable,"2")
'msgbox strStatus
''
'Public Function fnVerifyActiveStatus(objTable,strValue)
'	strActiveStatus = ""
'	intRows = objTable.RowCount
'	For intI = 1 To intRows
'		If objTable.GetCellData(intI,4) = strValue Then
'			strActiveStatus = objTable.GetCellData(intI,2)
'			Exit For
'		End If
'	Next
'	fnVerifyActiveStatus = strActiveStatus
'End Function
'
'
'
'Public Function fnSelectRequiredCheckBox(objTable,strValue)
'    blnFlag = False
'	intRows = objTable.RowCount
'	For intI = 1 To intRows
'		If objTable.GetCellData(intI,4) = strValue Then
'		   Setting.WebPackage("ReplayType") = 2
'		   objTable.ChildItem(intI,1,"Image",0).Click
'		   Setting.WebPackage("ReplayType") = 1
'		   blnFlag = True
'		   Exit For
'		End If
'	Next
'	fnSelectRequiredCheckBox = blnFlag
'End Function


'Set objParent = Browser("name:=.*Guidewire PolicyCenter").Page("title:=.*Guidewire PolicyCenter")
'''objParent.WebTable("html id:=gridview.*","class:=.*x-grid-table x-grid-with-col-lines x-grid-with-row-lines.*","index:=0").Highlight
''Set objTable = objParent.WebTable("html id:=gridview.*","class:=.*x-grid-table x-grid-with-col-lines x-grid-with-row-lines.*","index:=0")
''
''intRows = objTable.RowCount
'''intCols = objTable.ColumnCount(intRows)
''
'For intI = 1 to intRows	
'	If objTable.GetCellData(intI,4) = "2" Then
'	   Setting.WebPackage("ReplayType") = 2
'	   objTable.ChildItem(intI,1,"Image",0).Click
'	   Setting.WebPackage("ReplayType") = 1
'	End If
'Next    
''If intRows > 0 Then
'
'For intI = 1 to intRows	
'	For intJ = 1 to intCols
'		If Not IsNull(objTable.GetCellData(intI,intJ)) Then
'		   If objTable.GetCellData(intI,intJ) = "2" Then
'		   	 objTable.ChildItem(intI,1,"Image",0).Click 
'		   End If	
'		End If
'	Next
'Next
'	
'End If


'objParent.WebElement("html id:=.*AccountFile_LocationsScreen:ttlBar").Highlight

'Account File Locations
