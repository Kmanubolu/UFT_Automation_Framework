'==================================================================================================
'
'
'==================================================================================================
Public Function fnMenuNavigation(strMenuItem)
     blnFlag = False    
     fnMenuNavigation = True
     
     'Set OCommon = New OBJ_Common
     
    Select Case UCase(strMenuItem)
        Case UCase("DeskTop")
            Set objToClick = objDeskTop
            Set objToVerify = objDeskTopActions          
        Case UCase("Account")
            Set objToClick = objAccount
            Set objToVerify = objAccountActions        
        Case UCase("Policy")
            Set objToClick = objPolicy
            Set objToVerify = objPolicyActions    
        Case UCase("Contact")
            Set objToClick = objContact
            Set objToVerify = objContactActions
        Case UCase("Search")
            'Set objToClick = OCommon.objSearch
            'Set objToVerify = OCommon.objSearchPolicies    
            
            Set objToClick = objSearch
            Set objToVerify = objSearchPolicies    
        Case UCase("Team")    
            Set objToClick = objTeam
            Set objToVerify = objTeamMyGroupSummary    
        Case UCase("Administration")
            Set objToClick = objAdministration
            Set objToVerify = objAdminActions            
    End Select
    
   blnFlag = fnClickVerify(objToClick,objToVerify)
   If Not blnFlag Then
       fnMenuNavigation = False
    End If    
End Function

'==================================================================================================
'
'Function to click a specified item and verify whether specified item is available or not
'==================================================================================================
Public Function fnClickVerify(objToClick,objToVerify)
   	blnFlag = False
    If objToClick.Exist(Environment("SyncTimeOut")) Then
       Setting.WebPackage("ReplayType") = 2
       objToClick.Click
       Setting.WebPackage("ReplayType") = 1
       Wait 2
       If objToVerify.Exist(Environment("SyncTimeOut")) Then
          blnFlag = True
       End If              
   End If
   fnClickVerify = blnFlag
End Function

'==================================================================================================
'
'Function to click a specified item and verify whether specified item is available or not
'==================================================================================================
Public Function fnSpecificActionsItem(strClickingItem)
	fnSpecificActionsItem = True
	
	
    arrClickingItem = Split(strClickingItem,",")
    For intI = 0 To UBound(arrClickingItem)
      blnFlag = False
      Select Case UCase(arrClickingItem(intI))
        Case UCase("DeskTopActions")
             Set objToClick = objDeskTopActions    
             Set objToVerify = objNewAccount
        Case UCase("AccountActions")
             Set objToClick = objAccountActions    
             Set objToVerify = objNewSubmission             
        Case UCase("PolicyActions")
            Set objToClick = objPolicyActions
            Set objToVerify = objPolicyAccountFile            
        Case UCase("ContactActions")
            Set objToClick = objContactActions
            Set objToVerify = objNewAccount            
        Case UCase("AdminActions")
            Set objToClick = objAdminActions
            Set objToVerify = objAdminActionsNewUser            
        Case UCase("NewSubmission")
            Set objToClick = objNewSubmission
            Set objToVerify = objNewSubmissionsLbl
        Case UCase("NewAccount")
            Set objToClick = objNewAccount
            Set objToVerify = objAccountInfoLbl
        Case UCase("Move Policies to this Account")
            Set objToClick = objeleMovePoliciesToAccount
            Set objToVerify = objeleVerifyMovePoliciesToAccount
      End Select
      blnFlag = fnClickVerify(objToClick,objToVerify)
    Next
    If Not blnFlag Then
		fnSpecificActionsItem = False          
    End If              
End Function

'==================================================================================================
'
'Function to click a specified item and verify whether specified item is available or not
'==================================================================================================
Public Function fnSearchOrganization(strSearchCriteria,strValue)
    fnSearchOrganization = True
    If objOrganizationSearch.Exist(Environment.Value("SyncTimeOut")) Then
        objOrganizationSearch.Click
        If objOrganizationsTitle.Exist(Environment.Value("SyncTimeOut")) Then
             Select Case UCase(strSearchCriteria)
                 Case "ORGANIZATIONNAME"
                     objOrganizationsName.Object.Value = strValue
                 Case "ORGANIZATIONTYPE"     
                 Case "COUNTRY"
                 Case "CITY"
             End Select            
            objOrgSearchBtn.Click
            Wait(2)
            blnFlag = fnSelectItemInTable(strValue,objOrganizationTable)
            If not blnFlag Then
    		    fnSearchOrganization = False             	
            End If
            If objCreateAccount.Exist(Environment.Value("SyncTimeOut")) Then
            	    fnSearchOrganization = True
            End If
        Else
    		fnSearchOrganization = False 
    		Exit Function         
        End If
    Else
    	fnSearchOrganization = False    
    End If    
    
End Function

'==================================================================================================
'
'Function to click a specified item and verify whether specified item is available or not
'==================================================================================================
Public Function fnSelectItemInTable(strValue,objTable)
	fnSelectItemInTable = False
	If objTable.Exist(Environment.Value("SyncTimeOut")) Then
			For intI = 1 to objTable.Rowcount
			     	 strOrg = objTable.getcellData(intI,2)
			     	 If Trim(UCASE(strOrg)) = Trim(UCASE(strValue)) Then
			          	Set objSelect = objTable.ChildItem(intI,1,"WebElement",1)
			          	objSelect.HighLight
			          	objSelect.Click
			          	fnSelectItemInTable = True
			          	Exit Function
			      	End If     
			Next
	Else
			fnSelectItemInTable = False	
	End If
End Function

'==================================================================================================
'
'Function to click a specified item and verify whether specified item is available or not
'==================================================================================================
Function fnEnterWebTableListData(obj,strVal,intRow,intCol)
			fnEnterWebTableListData = False
			Set dd = obj.ChildItem(intRow,intCol,"WebElement",0) 
			dd.Click
			wait(1)
			BrowserPage().WebElement("html tag:=LI","innertext:=" & strVal).Click
			'set objListCollection = BrowserPage().WebElement("class:=x-boundlist x-boundlist-floating x-layer x-boundlist-default.*").Object.getElementsByTagName("LI")
			'set objListCollection = BrowserPage().WebElement("class:=x-boundlist x-boundlist-floating x-layer x-boundlist-default x-border-box x-boundlist-above").Object.getElementsByTagName("LI")
'			For each ListItem in objListCollection
'				If ListItem.innerText = strInput Then
'					ListItem.Click
'					Exit For
'				End If
'			Next
			wait(1)
			If dd.GetRoProperty("innertext") =  strVal Then
				fnEnterWebTableListData = True
			End If			
End Function


'==================================================================================================
'
'Function to click a specified item and verify whether specified item is available or not
'==================================================================================================
Function fnEnterWebTableEditData(obj,strVal,intRow,intCol,ClickCell)
			x = objParent.WebElement("index:=0","innertext:=" & ClickCell).GetROProperty("abs_X")
			y = objParent.WebElement("index:=0","innertext:=" & ClickCell).GetROProperty("abs_Y")
			objParent.WebElement("index:=0","innertext:=" & ClickCell).Click x,y
			Window("Internet Explorer").WinObject("Internet Explorer_Server").Type  micTab 
			wait(1)
			Window("Internet Explorer").WinObject("Internet Explorer_Server").Type  micTab 
			wait(1)
			Window("Internet Explorer").WinObject("Internet Explorer_Server").Type  micTab 

			wait(1)
			Set eObj = Description.Create
			eObj("micclass").Value = "WebEdit"
			eObj("class").Value = "x-form-field x-form-text x-form-focus x-field-form-focus x-field-default-form-focus"
			Set obj = objParent.ChildObjects(eObj)
			obj(0).object.value = strVal
			wait(1)
			fnEnterWebTableEditData = True
End Function

'==================================================================================================
'
'Function to click a specified item and verify whether specified item is available or not
'==================================================================================================
Function fnClickWebTableEditSearch(obj,strVal,intRow,intCol,ClickCell)
			fnClickWebTableEditSearch = True
			Set obj = BrowserPage()
			strLocation = ClickCell
			x = obj.WebElement("index:=0","innertext:=" & strLocation).GetROProperty("abs_X")
			y = obj.WebElement("index:=0","innertext:=" & strLocation).GetROProperty("abs_Y")
			obj.WebElement("index:=0","innertext:=" & strLocation).Click x,y
			Wait(2)
			Window("Internet Explorer").WinObject("Internet Explorer_Server").Type  micTab 


			wait(1)
			BrowserPage().WebElement("html id:=SubmissionWizard:LOBWizardStepGroup:LineWizardStepSet:WorkersCompCoverageConfigScreen:WorkersCompCoverageCV:WorkersCompClassesInputSet:WCCovEmpLV:" & intRow-1 &":ClassCode:SelectClassCode").Click
			If objeleClassCodeSearch.Exist(Environment.Value("SyncTimeOut")) Then
				fnClickWebTableEditSearch = True
			Else
				Exit Function
			End If
			strFlag = fnSelectItemInTable(strVal,objtblClassCodeSearch)
			If Not strFlag Then
					fnClickWebTableEditSearch = False
			End If
End Function

'==================================================================================================
'
'Function to select the Single or Multiple policy type
'==================================================================================================

Public Function fnPolicyType(strPolicies)
    blnFlag = False
    Select Case UCase(strPolicies)
    	Case "SINGLE"
    		Set objToClick = objSingle
    	Case "MULTIPLE"
    		Set objToClick = objMultiple
    End Select
    
    If objToClick.Exist Then
       objToClick.Click
       blnFlag = True
    End If
   fnPolicyType = blnFlag
   
End Function


'=================================================================================================='
'Function to select the radio button based on Yes or No input
'==================================================================================================
Public Function fnTableRadioBtnSelection(objTab,strYesNo)
    blnFlag = False
	Set objTable = objTab
	For intI = 1 To objTable.RowCount
	  Select Case UCase(strYesNo)  
	  	Case "YES"
	  		indx = 0
	  	Case "NO"
	  		indx = 1
	  End Select
	  If objTable.ChildItemCount(intI,2,"WebButton") > 0 Then
	  	objTable.ChildItem(intI,2,"WebButton",indx).Click	
	  	blnFlag = True
	  End If  
	Next
	fnTableRadioBtnSelection = blnFlag
End Function

'=================================================================================================='
'Function to select the radio button based on Question and Yes or No input
'==================================================================================================
Public Function fnSelectTableRadioBtn(objTab,strValue)
    blnFlag = False
    Set objTable = objTab
'    Select Case UCase(strYesNo)  
'          Case "YES"
'              indx = 0
'          Case "NO"
'              indx = 1
'      End Select
    For intI = 1 To objTable.RowCount
     If UCase(Trim(objTable.GetCellData(intI,1))) = UCase(Trim(strValue)) Then
         If objTable.ChildItemCount(intI,2,"WebButton") > 0 Then
            objTable.ChildItem(intI,2,"WebButton",0).Click    
            blnFlag = True
         End If  
     End If     
    Next
    fnSelectTableRadioBtn = blnFlag
End Function


'=================================================================================================='
'Function to select quick launch link on the left corner
'==================================================================================================

Public Function fnClickQuickLaunchLinks(objToClick,objToVerify)
	    blnFlag = False
		If objToClick.Exist(Environment("SyncTimeOut")) Then
	       Setting.WebPackage("ReplayType") = 2
	       objToClick.Click  'Click the specified link on left side corner
	       Setting.WebPackage("ReplayType") = 1
	       Wait 1
	       If objToVerify.Exist(Environment("SyncTimeOut")) Then
	       	  blnFlag = True
	       End If
	    Else
		   blnFlag = False				    
		End If 
		fnClickQuickLaunchLinks = blnFlag							    
	End Function


'=================================================================================================='
'Function to select check box of a specified item
'==================================================================================================

Public Function fnTableSelectSpecifiedCheckBox(objTab,strValue)
    blnFlag = False
	'Set objTab = objTable
	If objTab.Exist Then
	   intRows = objTab.RowCount
	   If intRows > 0  Then
	      intCols = objTab.Columncount(intRows)
			For intI = 1 To intRows
			  For intJ = 1 To intCols
			  	If InStr(objTab.GetCellData(intI,IntJ),strValue) > 0 Then
			  	   Setting.WebPackage("ReplayType") = 2	  	
			  	   objTab.ChildItem(intI,1,"Image",0).Click
			  	   Setting.WebPackage("ReplayType") = 1
			  	   blnFlag = True
				   Exit For 			  	   
			  	End If
			  Next
			  If blnFlag Then
			  	Exit For
			  End If	
			Next		
		End If	
	 End If 
   fnTableSelectSpecifiedCheckBox = blnFlag	
End Function


'=================================================================================================='
'Function to select the radio button based on Question and Yes or No input
'==================================================================================================
Public Function fnClickLeftPanelItems(objToClick,objToVerify)
    fnClickLeftPanelItems = False
	If objToClick.Exist(Environment("SyncTimeOut")) Then
       Setting.WebPackage("ReplayType") = 2
       objToClick.Click  'Click the Account which is on left side corner
       Setting.WebPackage("ReplayType") = 1
       If objToVerify.Exist(Environment("SyncTimeOut")) Then
       	  fnClickLeftPanelItems = True
       End If
    Else
	   fnClickLeftPanelItems = False				    
	End If 						    
End Function



'=================================================================================================='
'Function to select the radio button based on Question and Yes or No input
'==================================================================================================
Public Function fnReadTableData(obj,strCol,strSearchString)
    fnReadTableData = False
	If obj.Exist(Environment("SyncTimeOut")) Then
       For i = 1 to obj.RowCount
       	 If Instr(Trim(obj.GetCellData(i,strCol)),Trim(strSearchString))>0 Then
       			   fnReadTableData = True
       			   Exit Function
       	End If
       Next
	End If 						    
End Function

'=================================================================================================='
'Function to select the radio button based on Question and Yes or No input
'==================================================================================================
Public Function fnClickControlInTableBasedOnText(obj,strTextCol,strControlCol,strControl,strSearchString)
    fnClickControlInTableBasedOnText = False
	If obj.Exist(Environment("SyncTimeOut")) Then
       For i = 1 to obj.RowCount
       	 If Instr(Trim(obj.GetCellData(i,strTextCol)),Trim(strSearchString))>0 Then
 			'strCount = obj.ChildItemCount(i,strControlCol,strControl)    
 			Set objClick = obj.ChildItem(i,strControlCol,strControl,0) 
			Setting.WebPackage("ReplayType") = 2       	 			
 			objClick.Click
		    Setting.WebPackage("ReplayType") = 1
		    fnClickControlInTableBasedOnText = True
		   Exit Function
       	End If
       Next
	End If 						    
End Function

'=================================================================================================='
'Function to Select Actions items
'==================================================================================================
Public Function  fnSelectActionAccount(strActiontype,strValueToSelect,strValuetoVerify)
    blnFlag = False    
    Select Case strActionType
        Case "AccountActions"
            objAccountActions.Click
            blnFlag = True
        Case "PolicyActions"
    End Select
    
    blnFlag = False
    blnFlag = fnClickVerify(strValueToSelect,strValuetoVerify)
    
    fnSelectActionAccount = blnFlag        
End Function
'=================================================================================================='
'Function to Verify Table value
'==================================================================================================

Public Function fnTableverify(objTable,strValue)
	blnFlag = False
	intRows = objTable.RowCount
		If intRows > 0 Then
			   For intI = 1 To intRows					  
				  If Trim(objTable.GetCellData(intI,8)) = Trim(strValue) Then					  	
					 blnFlag = True
					 Exit For
				 End If
			   Next
		End If				
	fnTableverify = blnFlag
End Function


'============================================================================================================='
'Function to Click on Transaction#,Transaction Type,Policy and navigate to respective page based on the status 
'=============================================================================================================
Public Function fnTableClickSpecifiedItemNavigate(strSelectionType,strSelectionValue)
	blnFlag = False	
	Select Case strSelectionType
		Case "Policy"
			indx = 0
			intCol = 1
			intStatusCol = 3
		Case "Transaction" 
			indx = 1
			intCol = 2
			intStatusCol = 3
		Case "TransactionType"	
			indx = 1
			intCol = 6
			intStatusCol = 3
	End Select
	Set objParent = Browser("name:=.*Guidewire PolicyCenter.*").Page("title:=.*Guidewire PolicyCenter.*")
	Set objTable = objParent.WebTable("class:=.*x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","name:=WebTable","index:="&indx)
	intRows = objTable.RowCount
	If intRows > 0 Then
	   For intI = 1 to intRows     
		  If Trim(objTable.GetCellData(intI,intCol)) = Trim(strSelectionValue) Then
		  	 Environment("PolicyTransTypeStatus") = objTable.GetCellData(intI,intStatusCol)
			 Setting.WebPackage("ReplayType") = 2
			 objTable.ChildItem(intI,intCol,"WebElement",1).Click
			 Setting.WebPackage("ReplayType") = 1		 
			 Exit For		
		  End If
	   Next
	End If
	
	Select Case Environment("PolicyTransTypeStatus")
		Case "Draft"
		   Set objToVerify = objEleQualification
		Case "Quoted"
		   Set objToVerify = objEleQuoteTitle	
		Case "Canceled","In Force"
		   Set objToVerify = objEleSummary
	End Select
	
	'Verify whether page navigation to expected page is successful
	If objToVerify.Exist(Environment("SyncTimeOut")) Then  
	 	blnFlag = True
	End If
	
	fnTableClickSpecifiedItemNavigate = blnFlag
End Function


'=================================================================================================='
'Function to select the radio button based on Question and Yes or No input
'==================================================================================================
Public Function fnClickControlInTable(obj,strControlCol,strControlType,strSearchString,intIndex)
    fnClickControlInTable = False
	If obj.Exist(Environment("SyncTimeOut")) Then
			For i = 1 to obj.RowCount
				 If Instr(Trim(obj.GetCellData(i,strControlCol)),Trim(strSearchString))>0 Then
					Set objClick = obj.ChildItem(i,strControlCol,strControlType,intIndex) 
					Setting.WebPackage("ReplayType") = 2       	 			
					objClick.Click
			    	Setting.WebPackage("ReplayType") = 1
			    	fnClickControlInTable = True
			   		Exit Function
				End If
			Next
	End If 						    
End Function

