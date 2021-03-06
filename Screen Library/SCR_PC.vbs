'==================================================================================================
'
'
'==================================================================================================
Class SCR_Login

		Private oLogin

		Private Sub Class_initialize
					Set oLogin=New UI_Login
		End Sub

		Private Sub Class_Terminate
					Set oLogin=Nothing
		End Sub

		Public Function Login(strUserName,strPassword)
					'On Error Resume Next
					Environment.Value("CurrentUser") = Environment.Value("LoginUserName")
					Login =True
					If oLogin.m_Objects("edtUserName").Exist(Environment.Value("SyncTimeOut")) Then 
								Call fnInsertResult(Environment.Value("TestCaseID"),"Login Component","Guidewire Application should open successfully","Guidewire application opened successfully","PASS")
					Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Login Component","Guidewire Application should open successfully","Guidewire application opened successfully","FAIL")
								blnRunFlag=1
								Login=False
								Exit Function
					End If
					blnRetVal=SetPageValueDict(oLogin.m_Objects,oLogin,"edtUserName",strUserName)  
					If not blnRetVal Then
								blnRunFlag=1
								Login=False
								Exit Function
					End if
					If oLogin.m_Objects("edtPassword").Exist(Environment.Value("SyncTimeOut")) Then
								oLogin.m_Objects("edtPassword").SetSecure strPassword
					Else
								blnRunFlag=1
								Login=False
								Exit Function
					End If
					
					blnRetVal=SetPageValueDict(oLogin.m_Objects,oLogin,"eleLogin","Set")
					If not blnRetVal Then
								blnRunFlag=1
								Login=False
								Exit Function
					End if
					If oLogin.r_Objects("eleTabMenu").Exist(Environment.Value("SyncTimeOut")) Then 
								Call fnInsertResult(Environment.Value("TestCaseID"),"Login Component","User " & Chr(34) & strUserName & chr(34)&"should be logged in successfully","User " & Chr(34) & strUserName & Chr(34) & "has been logged successfully","PASS")
					Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Login Component","User " & Chr(34) & strUserName & chr(34)&"should be logged in successfully","User " & Chr(34) & strUserName & Chr(34) & "has been logged successfully","FAIL")
								blnRunFlag=1
								Login=False
								Exit Function
					End If
		End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_Logoff

			private oLogoff
			
			Private Sub Class_Initialize
						Set oLogoff = New UI_Logoff
			End Sub

			Private Sub Class_Terminate
						Set oLogoff = Nothing
			End Sub

			Public Function Logoff()
						'On Error Resume Next
						Logoff =True
						If oLogoff.m_Objects("elePereference").Exist(Environment.Value("SyncTimeOut")) Then
								oLogoff.m_Objects("elePereference").Click
						Else
								blnRunFlag=1
								Logoff = False
								Exit Function
						End If						
						If oLogoff.m_Objects("eleLogOut").Exist(Environment.Value("SyncTimeOut")) Then
								oLogoff.m_Objects("eleLogOut").Click
						Else
								blnRunFlag=1
								Logoff = False
								Exit Function
						End If
						If oLogoff.r_Objects("edtUserName").Exist(Environment.Value("SyncTimeOut")) Then
									
						ElseIf oLogoff.m_Objects("eleOk").Exist(Environment.Value("ObjectTimeOut")) Then
									oLogoff.m_Objects("eleOk").Click
						ElseIf oLogoff.m_Objects("btnOk").Exist(Environment.Value("ObjectTimeOut")) Then
									oLogoff.m_Objects("btnOk").Click
						Else
									Call fninsertResult(Environment.Value("TestCaseID"),"Logoff Component","User "&Chr(34) & Environment.Value("CurrentUser")& chr(34)& " should be logged off successfully","User was not logged off successfully","FAIL")
									blnRunFlag=1
									Logoff=False
									Exit Function
						End If
						Call fninsertResult(Environment.Value("TestCaseID"),"Logoff Component","User " & Chr(34) & Environment.Value("CurrentUser") & chr(34) & " should be logged off successfully","User has been logged off successfully","PASS")
						'Close all opened Browsers
						Call CloseOpenBrowser()
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_NewAccountSearch
			
			Private oNewAccountSearch

			Private Sub Class_Initialize
						Set oNewAccountSearch = New UI_NewAccountSearch
			End Sub

			Private Sub Class_Terminate
						Set oNewAccountSearch = Nothing
			End Sub
			
			Public Function NewAccountSearch()
						'On Error Resume Nex
						NewAccountSearch = True
						Wait 5			
						
						strFlag = fnMenuNavigation("Desktop")
						If Not strFlag Then
									NewAccountSearch = Falsep
									blnRunFlag = 1
									Exit Function
						End If
						strFlag = fnSpecificActionsItem("DeskTopActions,NewAccount")
						If Not strFlag Then
									NewAccountSearch = False
									blnRunFlag = 1
									Exit Function
						End If
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												If strClass1 = "edt" or strClass1 = "bli"  or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value																
												        blnRetVal = SetPageValueDict(oNewAccountSearch.m_Objects,oNewAccountSearch,ColName.Name, strValue)
														If Not blnRetVal Then
																	SearchAccount = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									'Click on Update
									'blnRetVal = SetPageValueDict(oCreateAccount.r_Objects,oCreateAccount,"eleUpdate","")
									If oNewAccountSearch.r_Objects("eleCreateAccount").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"NewAccountSearch Component","Create Account page should be displayed","Create Account page displayed", "PASS" )
									ElseIf oNewAccountSearch.r_Objects("edtCompanyName").Exist(Environment.Value("SyncTimeOut")) Then	
												Call fnInsertResult(Environment.Value("TestCaseID"),"NewAccountSearch Component","Search Address Book page should be displayed","Search Address Book page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"NewAccountSearch Component","Expected page should be displayed","Navigation to expected page was unsuccessful", "FAIL" )
												NewAccountSearch = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_CreateAccount
			
			Private oCreateAccount

			Private Sub Class_Initialize
						Set oCreateAccount = New UI_CreateAccount
			End Sub

			Private Sub Class_Terminate
						Set oCreateAccount = Nothing
			End Sub
			
			Public Function CreateAccount()
						'On Error Resume Nex
						CreateAccount=True

						For Each intKey in dictKeys'loop iterates number of test data records
									strAccountNum = "AccountNumber" & intkey & ""
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									'ContractInput
									strCreateAccount = 0
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value																
														If ColName.Name = "edtPostalCode" Then
															Wait(5)
															  blnRetVal = SetPageValueDict(oCreateAccount.m_Objects,oCreateAccount,ColName.Name, strValue)
														ElseIf ColName.Name = "edtOrganizationName" Then
															  blnRetVal = fnSearchOrganization("ORGANIZATIONNAME",strValue)
														Else
														      blnRetVal = SetPageValueDict(oCreateAccount.m_Objects,oCreateAccount,ColName.Name, strValue)
														End if
														If Not blnRetVal Then
																	CreateAccount = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									If oCreateAccount.r_Objects("eleAccountNumber").Exist(Environment.Value("SyncTimeOut")) Then 
												intAccountNumber = Split(ReadPageValueDict(oCreateAccount.r_Objects,oCreateAccount, "eleAccountNumber",1,1),"# ")
												Environment.Value(strAccountNum) = intAccountNumber(1)
												rsSheet("AccountNumber").Value = Environment.Value(strAccountNum)
												rsSheet.Update
												Call fnInsertResult(Environment.Value("TestCaseID"),"CreateAccount Component","Account should be saved and Account Number should be generated","Account has been saved and Account Number: "& Chr(34) & Environment.Value(strAccountNum) & Chr(34) & " has been generated", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"CreateAccount Component","Account should be saved and Account Number should be generated","Account Ticket has not been saved", "FAIL" )
												CreateAccount = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_SearchAccount
			
			Private oSearchAccount

			Private Sub Class_Initialize
						Set oSearchAccount = New UI_SearchAccount
			End Sub

			Private Sub Class_Terminate
						Set oSearchAccount = Nothing
			End Sub
			
			Public Function SearchAccount()
						'On Error Resume Nex
						blnFlag = False
						SearchAccount=True
						blnFlag = fnMenuNavigation("Search")
						If Not blnFlag Then
						   SearchAccount = False
						   blnRunFlag = 1
						   Exit Function
						End If		
						
						blnFlag = fnClickAccounts(oSearchAccount.m_Objects("eleSearchAccounts"),oSearchAccount.m_Objects("edtSearchEnterAccountNo"))						
						If Not blnFlag Then
						   SearchAccount = False
						   blnRunFlag = 1
						   Exit Function
						End If	
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
										strClass = UCASE(Left(ColName.Name, 3))
										strClass1 = Left(ColName.Name, 3)
										strDTColName = Mid(ColName.Name, 4)
										
										If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												strValue = ColName.Value																													
												If ColName.Name = "edtSearchEnterAccountNo" Then
													Wait(5)
													  Environment("AccountNumber") = strValue
													  blnFlag = fnSearchAccountNumber(strValue)
'														Else
'														      blnRetVal = SetPageValueDict(oSearchAccounts.m_Objects,oSearchAccounts,ColName.Name, strValue)
												End if
												If Not blnFlag Then
													SearchAccount = False
													blnRunFlag = 1
													Exit Function
												End If 
										End If
							Next
							If blnRunFlag <> 1 Then
								Call fnInsertResult(Environment.Value("TestCaseID"),"SearchAccounts Component","Searching Account in Search Page","Navigation to respective"& Environment("AccountNumber") &" Account Summary page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"SearchAccounts Component","Searching Account in Search Page","Navigation to respective Account Summary page was unsuccessful", "FAIL")
							End If									
						Next
			End Function
			
			'Function to Click Accounts WebElement in Search Page
			Public Function fnClickAccounts(objEleSearchAccounts,objEdtSearchEnterAccountNo)
			    blnFlag = False
				If objEleSearchAccounts.Exist(Environment("SyncTimeOut")) Then
			       Setting.WebPackage("ReplayType") = 2
			       objEleSearchAccounts.Click  'Click the Account which is on left side corner
			       Setting.WebPackage("ReplayType") = 1
			       If objEdtSearchEnterAccountNo.Exist(Environment("SyncTimeOut")) Then
			       	  blnFlag = True
			       End If
			    Else
				   blnFlag = False				    
				End If 
				fnClickAccounts = blnFlag							    
			End Function
			'Function to Enter Account Number in search page and Verify whether the expected Account Summary page is displayed or not
			Public Function fnSearchAccountNumber(strAccountNumber)
			    blnFlag = False
			    If objEdtSearchEnterAccountNo.Exist(Environment("SyncTimeOut")) Then
			       	  objEdtSearchEnterAccountNo.Object.Value = strAccountNumber ' Enter the account number to search
			       	  objLnkSearchBtn.Click 'Click Search button
			       	  Wait 1
			       	  If objEleSearchResultItem.Exist(Environment("SyncTimeOut")) Then
			       	  	objEleSearchResultItem.Click 'Click the account number link from the search results section
			       	  	Wait 2
			       	  	If objEleAccountSummaryAccountNo.Exist(Environment("SyncTimeOut")) Then
			       	  	   If inStr(objEleAccountSummaryAccountNo.GetROProperty("innertext"),strAccountNumber) > 0 Then 'Validate whether expected summary page is displayed or not
			       	  	   	  blnFlag = True
			       	  	   Else
							  blnFlag = False	       	  	   
			       	  	   End If	
			       	  	End If
			       	  End If
			    Else
					blnFlag = False    
			    End If        
				fnSearchAccountNumber = blnFlag            
			End Function
		
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_NewSubmissions
			
			Private oNewSubmissions

			Private Sub Class_Initialize
						Set oNewSubmissions = New UI_NewSubmissions
			End Sub

			Private Sub Class_Terminate
						Set oNewSubmissions = Nothing
			End Sub
			
			Public Function NewSubmissions()
						'On Error Resume Next
						blnFlag = False
						NewSubmissions=True
						
						blnFlag = fnSpecificActionsItem("AccountActions,NewSubmission")
						If Not blnFlag Then
						   NewSubmissions = False
						   blnRunFlag = 1
						   Exit Function
						End If	
						
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
							For Each ColName in rsSheet.Fields
								strClass = UCASE(Left(ColName.Name, 3))
								strClass1 = Left(ColName.Name, 3)
								strDTColName = Mid(ColName.Name, 4)
							
								If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" or strClass1 = "tbl" Then
									strValue = ColName.Value																
'									If ColName.Name = "lstQuoteType" or ColName.Name = "lstBaseState" Then
'										blnRetVal = fnSelectQuoteTypeBaseState(strDTColName,strValue)
									If ColName.Name = "tblLOB" Then
										blnRetVal = fnSelectLOB(strValue)
'									Else
'									    blnRetVal = SetPageValueDict(oNewSubmissions.m_Objects,oNewSubmissions,ColName.Name, strValue)
									End if
									If Not blnRetVal Then
										NewSubmissions = False
										blnRunFlag = 1
										Exit Function
									End If 
								End If
							Next
							If blnRunFlag <> 1 Then
								Call fnInsertResult(Environment.Value("TestCaseID"),"NewSubmissions Component","Entering all the required fields and selecting the Line of business in New Submissions page","Navigation to Qualification page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"NewSubmissions Component","Entering all the required fields and selecting the Line of business in New Submissions page","Navigation to Qualification page was unsuccessful", "FAIL")
							End If
					Next
			End Function
'			
			'Function to select the Line of Business
			Public Function fnSelectLOB(strValue)
				Set objTable = objTblLOB
				blnFlag = False
				If objTable.Exist(5) Then
					intRows = objTable.RowCount
					intCols = objTable.ColumnCount(intRows) 
					For intI = 1 to intRows
						For intJ = 1 To intCols
							If UCASE(Trim(objTable.GetCellData(intI,intJ))) = UCASE(Trim(strValue)) Then
								intRow = intI
								Exit For
								blnFlag = True
							End If
						Next
						If blnFlag Then
							Exit for	
						End If
					Next
					blnFlag = False
					For intK = 1 To intCols
						If UCASE(Trim(objTable.GetCellData(intRow,intK))) = UCASE(Trim("Select")) Then
							indx = intRow-1
							objTable.WebElement("innertext:=Select","html id:=NewSubmission:NewSubmissionScreen:ProductOffersDV:ProductSelectionLV.*","index:="&indx).Click
							blnFlag = True
							Exit For
						End If	
					Next
				End If
				blnFlag = False
				If objEleQualification.Exist(Environment("SyncTimeOut")) Then
				  	blnFlag = True
				End If
			   fnSelectLOB = blnFlag	
			End Function
			
			'Function to select QuoteType and Base State in New Submission Page
			Public Function fnSelectQuoteTypeBaseState(objType,objVal)
			    blnFlag = False
				Select Case UCase(objType)
					Case "QUOTETYPE"
						Set objToSelect = objQuoteType			
					Case "BASESTATE"
						Set objToSelect = objBaseState
				End Select
				If objToSelect.Exist Then
				   objToSelect.Object.Value = objVal
				   objToSelect.Click
				   blnFlag = True				   
				End If
			    fnSelectQuoteTypeBaseState = blnFlag 							    
			End Function


End Class
'==================================================================================================
'
'
'==================================================================================================
Class SCR_Qualifications
			
			Private oQualifications

			Private Sub Class_Initialize
						Set oQualifications = New UI_Qualifications
			End Sub

			Private Sub Class_Terminate
						Set oQualifications = Nothing
			End Sub
			
			Public Function Qualifications()
						'On Error Resume Nex
						Qualifications = True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									blnRetVal = True
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
														strValue = ColName.Value
														If ColName.Name = "Prequalification" Then
														   strValue = ColName.Value
														   blnRetVal = fnTableRadioBtnSelection(objTblPrequalification,strValue)
														ElseIf InStr(ColName.Name,"Question") > 0 Then														
										   				   strValue = ColName.Value
										   				   If Not IsNull(strValue) Then										   		
										   				   	  blnRetVal = fnSelectTableRadioBtn(objTblPrequalification,strValue)							   				   	  
														   End If														      				   
														End If
														If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
																blnRetVal = SetPageValueDict(oQualifications.m_Objects,oQualifications,ColName.Name, strValue)							
														End If
														If Not blnRetVal Then
																	Qualifications = False
																	blnRunFlag = 1
																	Exit Function
														End If 

									Next
									wait(3)
									If oQualifications.r_Objects("elePolicyInfo").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"Qualifications Component","Policy Info  page should be displayed","Policy Info page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"Qualifications Component","Policy Info  page should be displayed","Policy Info page not displayed", "FAIL" )
												Qualifications = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_PolicyInfo
			
		Private oPolicyInfo

		Private Sub Class_Initialize
					Set oPolicyInfo = New UI_PolicyInfo
		End Sub

		Private Sub Class_Terminate
					Set oPolicyInfo = Nothing
		End Sub
		
		Public Function PolicyInfo()
				    PolicyInfo = True
				For Each intKey in dictKeys'loop iterates number of test data records
						For Each ColName in rsSheet.Fields
							'arrKeys = oCreateAccount.m_Objects.Keys 'Gives all keys which are in the UI_MMContractInput Object definition class
							strClass = UCASE(Left(ColName.Name, 3))
							strClass1 = Left(ColName.Name, 3)
							strDTColName = Mid(ColName.Name, 4)
							
							If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
									strValue = ColName.Value	
									blnRetVal = SetPageValueDict(oPolicyInfo.m_Objects,oPolicyInfo,ColName.Name, strValue)							
									If Not blnRetVal Then
										PolicyInfo = False
										blnRunFlag = 1
										Exit Function
									End If 
							End If
						Next			
						wait(3)
						If oPolicyInfo.r_Objects("eleLocations").Exist(Environment.Value("SyncTimeOut")) Then 
									Call fnInsertResult(Environment.Value("TestCaseID"),"PolicyInfo Component","Locations page should be displayed","Locations page displayed", "PASS" )
						Else
									Call fnInsertResult(Environment.Value("TestCaseID"),"PolicyInfo Component","Locations page should be displayed","Locations page not displayed", "FAIL" )
									PolicyInfo  = False
									blnRunFlag = 1
									Exit Function
						End If
				Next			
		End Function

End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_Locations
			
			Private oLocations

			Private Sub Class_Initialize
						Set oLocations = New UI_Locations
			End Sub

			Private Sub Class_Terminate
						Set oLocations = Nothing
			End Sub
			
			Public Function Locations()
						'On Error Resume Nex
						Locations = True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
														blnRetVal = SetPageValueDict(oLocations.m_Objects,oLocations,ColName.Name, strValue)
														If Not blnRetVal Then
																	Locations = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									wait(3)
									If oLocations.r_Objects("eleWCCoverage").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Component","WC Coverage  page should be displayed","WC Coverage page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Component","WC Coverage  page should be displayed","WC Coverage page not displayed", "FAIL" )
												Locations = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_WCCoverage
			
			Private oWCCoverage

			Private Sub Class_Initialize
						Set oWCCoverage = New UI_WCCoverage
			End Sub

			Private Sub Class_Terminate
						Set oWCCoverage = Nothing
			End Sub
			
			Public Function WCCoverage()
						'On Error Resume Nex
						WCCoverage=True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
'									If Environment.Value("AccountNumber" & intkey) = "" Then
'												strAccountNumber =  rsSheet("AccountNumber")
'									Else
'												strAccountNumber =  Environment.Value("AccountNumber" & intkey)
'									End If
									
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
													
														If ColName.Name = "lstGoverningLaw" Then	
																	blnRetVal = fnEnterWebTableListData(oWCCoverage.m_Objects("tblAddClass"),strValue,1,2)
														ElseIf ColName.Name = "lstLocation" Then	
																	blnRetVal = fnEnterWebTableListData(oWCCoverage.m_Objects("tblAddClass"),strValue,1,3)
														ElseIf ColName.Name = "edtClassCode" Then
																blnRetVal = fnClickWebTableEditSearch(oWCCoverage.m_Objects("tblAddClass"),strValue,1,4,rsSheet("lstLocation"))
														ElseIf ColName.Name = "edtBasis" Then
																blnRetVal = fnEnterWebTableEditData(oWCCoverage.m_Objects("tblAddClass"),strValue,1,8,rsSheet("lstLocation"))
														Else
																blnRetVal = SetPageValueDict(oWCCoverage.m_Objects,oWCCoverage,ColName.Name, strValue)
														End If
														If Not blnRetVal Then
																	WCCoverage = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									If oWCCoverage.r_Objects("eleSupplimental").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"WCCoverage Component","Supplimental page should be displayed","Supplimental page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"WCCoverage Component","Supplimental page should be displayed","Supplimental page not displayed", "FAIL" )
												WCCoverage = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_Supplemental
			
			Private oSupplemental

			Private Sub Class_Initialize
						Set oSupplemental = New UI_Supplemental
			End Sub

			Private Sub Class_Terminate
						Set oSupplemental = Nothing
			End Sub
			
			Public Function Supplemental()
						'On Error Resume Nex
						Supplemental = True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												strValue = ColName.Value
												If ColName.Name = "SupplementalDefault" Then
												   strValue = ColName.Value
												   blnRetVal = fnTableRadioBtnSelection(objTblSupplemental,strValue)
												ElseIf InStr(ColName.Name,"Question") > 0 Then														
								   				   strValue = ColName.Value
								   				   If Not IsNull(strValue) Then										   		
								   				   	  blnRetVal = fnSelectTableRadioBtn(objTblSupplemental,strValue)							   				   	  
												   End If														      				   
												End If
											
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
														blnRetVal = SetPageValueDict(oSupplemental.m_Objects,oSupplemental,ColName.Name, strValue)
														If Not blnRetVal Then
															Supplemental = False
															blnRunFlag = 1
															Exit Function
														End If 
												End If
									Next
									wait(3)
									If oSupplemental.r_Objects("eleWCOptions").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"Supplemental Component","WC Options page should be displayed","WC Options page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"Supplemental Component","WC Options page should be displayed","WC Options page not displayed", "FAIL" )
												Supplemental = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_WCOptions
			
			Private oWCOptions

			Private Sub Class_Initialize
						Set oWCOptions = New UI_WCOptions
			End Sub

			Private Sub Class_Terminate
						Set oWCOptions = Nothing
			End Sub
			
			Public Function WCOptions()
						'On Error Resume Nex
						WCOptions = True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
														blnRetVal = SetPageValueDict(oWCOptions.m_Objects,oWCOptions,ColName.Name, strValue)
														If Not blnRetVal Then
																	WCOptions = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									If oWCOptions.r_Objects("eleRiskAnalysis").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"WCOptions Component","Risk Analysis Page should be displayed","Risk Analysis page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"WCOptions Component","Risk Analysis Page should be displayed","Risk Analysis page not displayed", "FAIL" )
												WCOptions = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_RiskAnalysis
			
			Private oRiskAnalysis

			Private Sub Class_Initialize
						Set oRiskAnalysis = New UI_RiskAnalysis
			End Sub

			Private Sub Class_Terminate
						Set oRiskAnalysis = Nothing
			End Sub
			
			Public Function RiskAnalysis()
						'On Error Resume Nex
						RiskAnalysis  =True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
								
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
													
														blnRetVal = SetPageValueDict(oRiskAnalysis.m_Objects,oRiskAnalysis,ColName.Name, strValue)
														If Not blnRetVal Then
																	RiskAnalysis = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									If oRiskAnalysis.r_Objects("elePolicyReview").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"RiskAnalysis Component","Policy Review page should be displayed","Policy Review page displayed","PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"RiskAnalysis Component","Policy Review page should be displayed","Policy Review page not displayed","FAIL" )
												RiskAnalysis = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_PolicyReview
			
			Private oPolicyReview

			Private Sub Class_Initialize
						Set oPolicyReview = New UI_PolicyReview
			End Sub

			Private Sub Class_Terminate
						Set oPolicyReview = Nothing
			End Sub
			
			Public Function PolicyReview()
						'On Error Resume Nex
						PolicyReview = True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
														blnRetVal = SetPageValueDict(oPolicyReview.m_Objects,oPolicyReview,ColName.Name, strValue)
														If Not blnRetVal Then
																	PolicyReview = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									wait(3)
									If oPolicyReview.r_Objects("eleQuote").Exist(Environment.Value("SyncTimeOut")) Then
												If oPolicyReview.m_Objects("eleQuoteAlert").Exist(5) Then
													If oPolicyReview.m_Objects("eleRiskAnalysisTree").Exist(60)  Then
														Setting.WebPackage("ReplayType") = 2	
														oPolicyReview.m_Objects("eleRiskAnalysisTree").Click
														Setting.WebPackage("ReplayType") = 1
													End If
													If oPolicyReview.m_Objects("tblRiskAnalysisIssuesTable").Exist(60) Then
														rowCnt = oPolicyReview.m_Objects("tblRiskAnalysisIssuesTable").RowCount
														For Iterator = 2 To rowCnt Step 1
														   Setting.WebPackage("ReplayType") = 2	
															oPolicyReview.m_Objects("tblRiskAnalysisIssuesTable").ChildItem(Iterator,1,"Image",0).Click
															Setting.WebPackage("ReplayType") = 1
															Wait 1
														Next
														If oPolicyReview.m_Objects("eleRiskAnalysisApprove").Exist(60) Then
															oPolicyReview.m_Objects("eleRiskAnalysisApprove").Click
															Wait(3)
															oPolicyReview.m_Objects("eleRiskAnalysisOK").RefreshObject
															Wait(5)
															If oPolicyReview.m_Objects("eleRiskAnalysisOK").Exist(60) Then
																oPolicyReview.m_Objects("eleRiskAnalysisOK").Click
																Wait(3)
															End If
														End If
														If oPolicyReview.m_Objects("eleTreeQuote").Exist(60) Then
															Setting.WebPackage("ReplayType") = 2
															oPolicyReview.m_Objects("eleTreeQuote").Click
															Setting.WebPackage("ReplayType") = 1
														End If														
													End If
																
												End If
												
												Call fnInsertResult(Environment.Value("TestCaseID"),"PolicyReview Component","Quote page should be displayed","Quote page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"PolicyReview Component","Quote page should be displayed","Quote page not displayed", "FAIL" )
												PolicyReview = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_Quote
			
			Private oQuote

			Private Sub Class_Initialize
						Set oQuote = New UI_Quote
			End Sub

			Private Sub Class_Terminate
						Set oQuote = Nothing
			End Sub
			
			Public Function Quote()
						'On Error Resume Nex
						Quote = True
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "bli" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
														blnRetVal = SetPageValueDict(oQuote.m_Objects,oQuote,ColName.Name, strValue)
														If Not blnRetVal Then
																	Quote = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									
									If oQuote.r_Objects("eleSubmissionBound").Exist(Environment.Value("SyncTimeOut")) Then 
												intPolicyNumber = ReadPageValueDict(oQuote.r_Objects,oQuote, "tblSubmissionInfo",4,1)
												rsSheet("PolicyNumber").Value = intPolicyNumber
												rsSheet.Update
												Call fnInsertResult(Environment.Value("TestCaseID"),"Quote Component","Submission Bound page should be displayed","Submission Bound page displayed", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"Quote Component","Submission Bound page should be displayed","Submission Bound page not displayed", "FAIL" )
												Quote = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_UpdateAccount
			
			Private oUpdateAccount

			Private Sub Class_Initialize
						Set oUpdateAccount = New UI_UpdateAccount
			End Sub

			Private Sub Class_Terminate
						Set oUpdateAccount = Nothing
			End Sub
			
			Public Function UpdateAccount()
						'On Error Resume Nex
						UpdateAccount=True
						
						For Each intKey in dictKeys'loop iterates number of test data records
									strAccountNum = "AccountNumber" & intkey & ""
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									'ContractInput
									 For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value																
														If ColName.Name = "edtPostalCode" Then
															Wait(5)
															  blnRetVal = SetPageValueDict(oUpdateAccount.m_Objects,oUpdateAccount,ColName.Name, strValue)
'														ElseIf ColName.Name = "edtOrganizationName" Then
'															  blnRetVal = fnSearchOrganization("ORGANIZATIONNAME",strValue)
														Else
														      blnRetVal = SetPageValueDict(oUpdateAccount.m_Objects,oUpdateAccount,ColName.Name, strValue)
														End if
														If Not blnRetVal Then
															UpdateAccount = False
															blnRunFlag = 1
															Exit Function
														End If 
												End If
									Next
									If oUpdateAccount.r_Objects("eleAccountSummary").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"UpdateAccount Component","Changes should be updated","Updation is Successful", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"UpdateAccount Component","Changes should be updated","Updation is Unsuccessful", "FAIL" )
												UpdateAccount = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_AddRelatedAccounts
			
			Private oAddRelatedAccounts

			Private Sub Class_Initialize
						Set oAddRelatedAccounts = New UI_AddRelatedAccounts
			End Sub

			Private Sub Class_Terminate
						Set oAddRelatedAccounts= Nothing
			End Sub
			
			Public Function AddRelatedAccounts()
						'On Error Resume Nex
						AddRelatedAccounts=True
						
						blnRetVal = fnClickQuickLaunchLinks(objEleRelatedAccounts,objEleAccountFileRelatedAccounts)
						If Not blnRetVal Then
							AddRelatedAccounts = False
							blnRunFlag = 1
							Exit Function
						End If
						For Each intKey in dictKeys'loop iterates number of test data records
									strAccountNum = "AccountNumber" & intkey & ""
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									'ContractInput
									 For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												   strValue = ColName.Value																
												   blnRetVal = SetPageValueDict(oAddRelatedAccounts.m_Objects,oAddRelatedAccounts,ColName.Name, strValue)
												End if
												If Not blnRetVal Then
													AddRelatedAccounts = False
													blnRunFlag = 1
													Exit Function
												End If
									
									Next
									If oAddRelatedAccounts.r_Objects("eleAccountFileRelatedAccounts").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"Add Related Accounts Component","Account Number need to be related","Updation is Successful", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"Add Related Accounts Component","Account Number need to be related","Updation is Unsuccessful", "FAIL" )
												AddRelatedAccounts = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_RemoveRelatedAccounts
			
			Private oRemoveRelatedAccounts

			Private Sub Class_Initialize
						Set oRemoveRelatedAccounts = New UI_RemoveRelatedAccounts
			End Sub

			Private Sub Class_Terminate
						Set oRemoveRelatedAccounts= Nothing
			End Sub
			
			Public Function RemoveRelatedAccounts()
						'On Error Resume Nex
						RemoveRelatedAccounts=True
						If Not objEleAccountFileRelatedAccounts.Exist(5) Then
						   blnRetVal = fnClickQuickLaunchLinks(objEleRelatedAccounts,objEleAccountFileRelatedAccounts)
						   If Not blnRetVal Then
							  RemoveRelatedAccounts = False
							  blnRunFlag = 1
							  Exit Function
						   End If
						End If
						
						For Each intKey in dictKeys'loop iterates number of test data records
									strAccountNum = "AccountNumber" & intkey & ""
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									'ContractInput
									 For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												strValue = ColName.Value
												If ColName.Name = "tblRemoveRelatedAccounts"Then
												   Environment("AccountNumber") = strValue
												   blnRetVal = fnTableSelectSpecifiedCheckBox(objTblRemoveRelatedAccounts,strValue) 
												End If													
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												   blnRetVal = SetPageValueDict(oRemoveRelatedAccounts.m_Objects,oRemoveRelatedAccounts,ColName.Name, strValue)
												End if													
												If Not blnRetVal Then
													RemoveRelatedAccounts = False
													blnRunFlag = 1
													Exit Function
												End If
									
									Next
									If oRemoveRelatedAccounts.r_Objects("eleAccountFileRelatedAccounts").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"RemoveRelationshipAccount Component","Remove Relationship from Account","Relationship removal for "&Environment("AccountNumber")&" was Successful", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"RemoveRelationshipAccount Component","Remove Relationship from Account","Relationship removal for "&Environment("AccountNumber")&" was Unsuccessful", "FAIL" )
												RemoveRelatedAccounts = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_CreateLocation
			
			Private oCreateLocation

			Private Sub Class_Initialize
						Set oCreateLocation = New UI_CreateLocation
			End Sub

			Private Sub Class_Terminate
						Set oCreateLocation = Nothing
			End Sub
			
			Public Function CreateLocation()
						'On Error Resume Nex
						CreateLocation = True
						blnFlag = fnClickLeftPanelItems(oCreateLocation.m_Objects("eleLocation"),oCreateLocation.r_Objects("eleAccountFileLocations"))						
						If Not blnFlag Then
						   CreateLocation = False
						   blnRunFlag = 1
						   Exit Function
						End If	
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												
												If strClass1 = "edt" or strClass1 = "bli" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
														If ColName.Name = "edtPostalCode" Then
																Wait(5)
																blnRetVal = SetPageValueDict(oCreateLocation.m_Objects,oCreateLocation,ColName.Name, strValue)
														ElseIf ColName.Name = "btnNonSpecificLocation" Then
															If strValue = "NO" Then
																blnRetVal = SetPageValueDict(oCreateLocation.m_Objects,oCreateLocation,"btnNonSpecificLocationNO", strValue)
															ElseIf strValue = "YES" Then
																blnRetVal = SetPageValueDict(oCreateLocation.m_Objects,oCreateLocation,"btnNonSpecificLocationYES", strValue)
														    End If
														Else
																blnRetVal = SetPageValueDict(oCreateLocation.m_Objects,oCreateLocation,ColName.Name, strValue)
														End If
														If Not blnRetVal Then
																	CreateLocation = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									
									If oCreateLocation.r_Objects("eleAccountFileLocations").Exist(Environment.Value("SyncTimeOut")) Then 
												strFlag = fnReadTableData(oCreateLocation.r_Objects("tblLocationsList"),8,rsSheet("edtAddressLine1") & ", "	& rsSheet("edtCity"))
												If strFlag Then
														Call fnInsertResult(Environment.Value("TestCaseID"),"CreateLocation Component","Location should be created","Location created", "PASS" )
												Else
														Call fnInsertResult(Environment.Value("TestCaseID"),"CreateLocation Component","Location should be created","Location not created", "FAIL" )
												End If
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"CreateLocation Component","Location should be created","Location not created", "FAIL" )
												CreateLocation = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_DeleteLocation
			
			Private oDeleteLocation

			Private Sub Class_Initialize
						Set oDeleteLocation = New UI_DeleteLocation
			End Sub

			Private Sub Class_Terminate
						Set oDeleteLocation = Nothing
			End Sub
			
			Public Function DeleteLocation()
						'On Error Resume Nex
						DeleteLocation = True
						blnFlag = fnClickLeftPanelItems(oDeleteLocation.m_Objects("eleLocation"),oDeleteLocation.r_Objects("eleAccountFileLocations"))						
						If Not blnFlag Then
						   DeleteLocation = False
						   blnRunFlag = 1
						   Exit Function
						End If	
						For Each intKey in dictKeys'loop iterates number of test data records
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												
												If strClass1 = "edt" or strClass1 = "bli" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
														strValue = ColName.Value
														If ColName.Name = "eleLocationToRemove" Then												
																blnRetVal = fnClickControlInTableBasedOnText(oDeleteLocation.r_Objects("tblLocationsList"),8,1,"Image",strValue)
														Else
																blnRetVal = SetPageValueDict(oDeleteLocation.m_Objects,oDeleteLocation,ColName.Name, strValue)
														End If
														If Not blnRetVal Then
																	CreateLocation = False
																	blnRunFlag = 1
																	Exit Function
														End If 
												End If
									Next
									
									If oDeleteLocation.r_Objects("eleAccountFileLocations").Exist(Environment.Value("SyncTimeOut")) Then 
												strFlag = fnReadTableData(oDeleteLocation.r_Objects("tblLocationsList"),8,rsSheet("eleLocationToRemove"))
												If not strFlag Then
														Call fnInsertResult(Environment.Value("TestCaseID"),"DeleteLocation Component","Location should be deleted","Location deleted", "PASS" )
												Else
														Call fnInsertResult(Environment.Value("TestCaseID"),"DeleteLocation Component","Location should be deleted","Location not deleted", "FAIL" )
												End If
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"DeleteLocation Component","Location should be deleted","Location not deleted", "FAIL" )
												CreateLocation = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================

Class SCR_SearchAddressBook

			Private oSearchAddressBook

			Private Sub Class_Initialize
						Set oSearchAddressBook = New UI_SearchAddressBook
			End Sub

			Private Sub Class_Terminate
						Set oSearchAddressBook= Nothing
			End Sub
			
			Public Function SearchAddressBook()			
						SearchAddressBook=True		
						For Each intKey in dictKeys'loop iterates number of test data records
								rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
								For Each ColName in rsSheet.Fields
									strClass = UCASE(Left(ColName.Name, 3))
									strClass1 = Left(ColName.Name, 3)
									strDTColName = Mid(ColName.Name, 4)
									
									If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" or strClass1 = "tbl" Then
									    strValue = ColName.Value
										If ColName.Name = "lstType" Then										    
									 		blnRetVal = SetPageValueDict(oSearchAddressBook.m_Objects,oSearchAddressBook,ColName.Name, strValue)
									 		Wait 3
									 	ElseIf ColName.Name = "tblAddressBookSelect" Then 
									 		Wait 3
											blnRetVal = fnTableSearchAddressBookBtnSelect(objTblSearchAddressBookSelect,strValue)			 		
									 	Else
											blnRetVal = SetPageValueDict(oSearchAddressBook.m_Objects,oSearchAddressBook,ColName.Name, strValue)
										End If	
									 		
										If Not blnRetVal Then
											SearchAddressBook = False
											blnRunFlag = 1
											Exit Function
										End If 	
								  End if																	
								Next
								If oSearchAddressBook.r_Objects("eleCreateAccountTitle").Exist(Environment.Value("SyncTimeOut")) Then 
											Call fnInsertResult(Environment.Value("TestCaseID"),"Search Address Book Component","Navigation to Create Account page should be successful","Navigation to Create Account page was Successful", "PASS" )
								Else
											Call fnInsertResult(Environment.Value("TestCaseID"),"Search Address Book Component","Navigation to Create Account page should be successful","Navigation to Create Account page was Unsuccessful", "FAIL" )
											SearchAddressBook = False
											blnRunFlag = 1
											Exit Function
								End If									
						Next
			End Function
			
'			Function to Select the select button from the Address Book Search results
			Public Function fnTableSearchAddressBookBtnSelect(objTab,strValue)
				Set objTable = objTab
				blnFlag = False
				If objTable.Exist Then
					intRows = objTable.RowCount
					intCols = objTable.ColumnCount(intRows) 
					For intI = 1 to intRows
						For intJ = 1 To intCols
							If UCASE(Trim(objTable.GetCellData(intI,intJ))) = UCASE(Trim(strValue)) Then
								intRow = intI
								blnFlag = True
								Exit For							
							End If
						Next
						If blnFlag Then
							Exit for	
						End If
					Next
					blnFlag = False
					For intK = 1 To intCols
						If UCASE(Trim(objTable.GetCellData(intRow,intK))) = UCASE(Trim("Select")) Then
							indx = intRow-1
							objTable.WebElement("innertext:=Select","html id:=.*ContactSearchResultsLV:0:_Select","index:="&indx).Click
							blnFlag = True
							Exit For
						End If	
					Next
				End If
			   fnTableSearchAddressBookBtnSelect = blnFlag	
			End Function	
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_SearchRelatedAccount
			
			Private oSearchRelatedAccount

			Private Sub Class_Initialize
						Set oSearchRelatedAccount = New UI_SearchRelatedAccount
			End Sub

			Private Sub Class_Terminate
						Set oSearchRelatedAccount= Nothing
			End Sub
			
			Public Function SearchRelatedAccount()
						'On Error Resume Next
						SearchRelatedAccount=True
						
						'Click on Related Accounts
						If Not objEleAccountFileRelatedAccounts.Exist(Environment("SyncTimeOut")) Then
							blnRetVal = fnClickQuickLaunchLinks(objEleRelatedAccounts,objEleAccountFileRelatedAccounts)
							If Not blnRetVal Then
								SearchRelatedAccount = False
								blnRunFlag = 1
								Exit Function
							End If
						End If	
						
						
						For Each intKey in dictKeys'loop iterates number of test data records
									
									rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"
									'ContractInput
									 For Each ColName in rsSheet.Fields
												strClass = UCASE(Left(ColName.Name, 3))
												strClass1 = Left(ColName.Name, 3)
												strDTColName = Mid(ColName.Name, 4)
												strValue = ColName.Value
												If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												  If ColName.Name = "edtSearchRelatedAccountAccountNo" Then
												  	Wait 10 
												  	blnRetVal = SetPageValueDict(oSearchRelatedAccount.m_Objects,oSearchRelatedAccount,ColName.Name,strValue)
												  Else	
												  	blnRetVal = SetPageValueDict(oSearchRelatedAccount.m_Objects,oSearchRelatedAccount,ColName.Name,strValue)
												  End If
												  
												   If Not blnRetVal Then
														SearchRelatedAccount = False
														blnRunFlag = 1
														Exit Function
													End If 
												End if
												If ColName.Name = "tblRelatedAccountSelect" Then
												   blnRetVal = fnTableSearchRelatedAccountSelect(objTblRelatedAccountSelect,strValue)
												   If Not blnRetVal Then
														SearchRelatedAccount = False
														blnRunFlag = 1
														Exit Function
												   End If 
												End If																								
									Next
									If oSearchRelatedAccount.r_Objects("eleAccountRelationship").Exist(Environment.Value("SyncTimeOut")) Then 
												Call fnInsertResult(Environment.Value("TestCaseID"),"Search Related Account Component","Navigation to Account Relationship page","Navigation to Account Relationship page is Successful", "PASS" )
									Else
												Call fnInsertResult(Environment.Value("TestCaseID"),"Search Related Account Component","Navigation to Account Relationship page","Navigation to Account Relationship page is Unsuccessful", "FAIL" )
												SearchRelatedAccount = False
												blnRunFlag = 1
												Exit Function
									End If
						Next
			End Function
			
			'Select the Account number from the search result
			Public Function fnTableSearchRelatedAccountSelect(objTable,strAccountNo)
			   blnFlag = False
			   Set objTab = objTable
			   If Trim(objTab.GetCellData(1,2)) = Trim(strAccountNo) Then
			      objTab.ChildItem(1,1,"WebElement",1).Click
			      blnFlag = True
			   End If
			   fnTableSearchRelatedAccountSelect = blnFlag
			End Function

		
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_SearchAccountPolicyMove
			
			Private oSearchAccountPolicyMove

			Private Sub Class_Initialize
						Set oSearchAccountPolicyMove = New UI_SearchAccountPolicyMove
			End Sub

			Private Sub Class_Terminate
						Set oSearchAccountPolicyMove = Nothing
			End Sub
			
			Public Function SearchAccountPolicyMove()
						'On Error Resume Nex
						blnFlag = False
						SearchAccountPolicyMove = True
						
						blnFlag = fnSpecificActionsItem("AccountActions,Move Policies to this Account")
						If Not blnFlag Then
						   SearchAccountPolicyMove = False
						   blnRunFlag = 1
						   Exit Function
						End If
						
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
										strClass = UCASE(Left(ColName.Name, 3))
										strClass1 = Left(ColName.Name, 3)
										strDTColName = Mid(ColName.Name, 4)
										
										If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												strValue = ColName.Value																													
												If ColName.Name = "tblSelectAccount" Then
														blnRetVal = fnSearchAccountNumberForMove(oSearchAccountPolicyMove.m_Objects("tblSelectAccount"),rsSheet("edtSearchEnterAccountNo"))
												Else
														blnRetVal = SetPageValueDict(oSearchAccountPolicyMove.m_Objects,oSearchAccountPolicyMove,ColName.Name, strValue)
												End if
												If Not blnFlag Then
													SearchAccountPolicyMove = False
													blnRunFlag = 1
													Exit Function
												End If 
										End If
							Next
							If oSearchAccountPolicyMove.r_Objects("eleMovepoliciesSelection").Exist(Environment.Value("SyncTimeOut")) Then 
								Call fnInsertResult(Environment.Value("TestCaseID"),"SearchAccounts Component","Searching Account in Search Page","Navigation to respective"& Environment("AccountNumber") &" Account Summary page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"SearchAccounts Component","Searching Account in Search Page","Navigation to respective Account Summary page was unsuccessful", "FAIL")
								SearchAccountPolicyMove = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
			Public Function fnSearchAccountNumberForMove(obj,strAccountNumber)
			    	  blnFlag = False
			       	  If obj.Exist(Environment("SyncTimeOut")) Then
			       	  	Setting.WebPackage("ReplayType") = 2       	 			
			       	  	Set ob = obj.ChildItem(1,1,"WebElement",1)
			       	  	ob.Click
			       	  	Setting.WebPackage("ReplayType") = 1      	 			
						blnFlag = True
			       	 End If
				fnSearchAccountNumberForMove = blnFlag            
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_AccountPolicyMove
			
			Private oAccountPolicyMove

			Private Sub Class_Initialize
						Set oAccountPolicyMove = New UI_AccountPolicyMove
			End Sub

			Private Sub Class_Terminate
						Set oAccountPolicyMove = Nothing
			End Sub
			
			Public Function AccountPolicyMove()
						'On Error Resume Nex
						blnFlag = False
						AccountPolicyMove = True
						
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
										strClass = UCASE(Left(ColName.Name, 3))
										strClass1 = Left(ColName.Name, 3)
										strDTColName = Mid(ColName.Name, 4)
										
										If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												strValue = ColName.Value																													
											
												If ColName.Name = "tblPolictSelection" Then
														'Wait(5)
														blnRetVal = fnSelectPoliciesToMove(oAccountPolicyMove.m_Objects("tblPolictSelection"),rsSheet("tblPolictSelection"))
												Else
														blnRetVal = SetPageValueDict(oAccountPolicyMove.m_Objects,oAccountPolicyMove,ColName.Name, strValue)
												End if
												If Not blnRetVal Then
													AccountPolicyMove = False
													blnRunFlag = 1
													Exit Function
												End If 
										End If
							Next
							If oAccountPolicyMove.r_Objects("eleAccountFileSummary").Exist(Environment.Value("SyncTimeOut")) Then 
								Call fnInsertResult(Environment.Value("TestCaseID"),"SearchAccounts Component","Searching Account in Search Page","Navigation to respective"& Environment("AccountNumber") &" Account Summary page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"SearchAccounts Component","Searching Account in Search Page","Navigation to respective Account Summary page was unsuccessful", "FAIL")
								AccountPolicyMove = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
			Public Function fnSelectPoliciesToMove(obj,strPplicyNumber)
			    	  blnFlag = False
			       	  If obj.Exist(Environment("SyncTimeOut")) Then
				       	  	For i = 1 To obj.RowCount
				       	  			If obj.GetCellData(i,2) = strPplicyNumber Then
				       	  					Setting.WebPackage("ReplayType") = 2       	 			
				       	  					Set ob = obj.ChildItem(i,1,"Image",0)
				       	  					ob.Click
				       	  					Setting.WebPackage("ReplayType") = 1
				       	  					blnFlag = True
				       	  					Exit For
				       	  			End If
							Next
			       	 End If
				fnSelectPoliciesToMove = blnFlag            
			End Function		
End Class

'==================================================================================================
'
'
'==================================================================================================

Class SCR_AddRoleParticipants

			Private oAddRoleParticipants

			Private Sub Class_Initialize
						Set oAddRoleParticipants = New UI_AddRoleParticipants
			End Sub

			Private Sub Class_Terminate
						Set oAddRoleParticipants= Nothing
			End Sub
			
			Public Function AddRoleParticipants()			
						AddRoleParticipants=True			
   						blnFlag = fnClickQuickLaunchLinks(objEleParticipants,objEleParticipantsTitle)			
   		  	
						If Not blnFlag Then
						   AddRoleParticipants = False
						   blnRunFlag = 1
						   Exit Function
						End If
						For Each intKey in dictKeys'loop iterates number of test data records
								rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
								For Each ColName in rsSheet.Fields
									strClass = UCASE(Left(ColName.Name, 3))
									strClass1 = Left(ColName.Name, 3)
									strDTColName = Mid(ColName.Name, 4)
									
									If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
										strValue = ColName.Value
										If ColName.Name = "eleAssignGroupCol" Then
											Wait(10)
											blnRetVal = SetPageValueDict(oAddRoleParticipants.m_Objects,oAddRoleParticipants,ColName.Name, strValue)
										Else
											blnRetVal = SetPageValueDict(oAddRoleParticipants.m_Objects,oAddRoleParticipants,ColName.Name, strValue)
										End if
										If Not blnFlag Then
											AddRoleParticipants = False
											blnRunFlag = 1
											Exit Function
										End If 
									End If
								Next
							If blnRunFlag <> 1 Then							  
								Call fnInsertResult(Environment.Value("TestCaseID"),"Add Role Participants Component","Adding Role to Participants","Add Roles to Participants was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Add Role Participants Component","Adding Role to Participants","Add Roles to Participants was unsuccessful", "FAIL")
							End If									
						Next
			End Function
						
				
End Class


'==================================================================================================
'
'
'==================================================================================================

Class SCR_RemoveRoleParticipants

			Private oRemoveRoleParticipants

			Private Sub Class_Initialize
						Set oRemoveRoleParticipants = New UI_RemoveRoleParticipants
			End Sub

			Private Sub Class_Terminate
						Set oRemoveRoleParticipants= Nothing
			End Sub
			
			Public Function RemoveRoleParticipants()			
					
							For Each intKey in dictKeys'loop iterates number of test data records
								rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
								For Each ColName in rsSheet.Fields
									strClass = UCASE(Left(ColName.Name, 3))
									strClass1 = Left(ColName.Name, 3)
									strDTColName = Mid(ColName.Name, 4)
									
									If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
										strValue = ColName.Value
										If ColName.Name = "eleRemoveSelectCheck" Then 
											Wait 5
										 	blnRetVal = fnTableSelectSpecifiedCheckBox(objTblRemoveRoleParticipants,strValue)										
										ElseIf ColName.Name = "eleRemove" or ColName.Name = "eleUpdate" Then
											Wait 3
											blnRetVal = SetPageValueDict(oRemoveRoleParticipants.m_Objects,oRemoveRoleParticipants,ColName.Name,strValue)
										Else	
											blnRetVal = SetPageValueDict(oRemoveRoleParticipants.m_Objects,oRemoveRoleParticipants,ColName.Name,strValue)
										End if
										If Not blnRetVal Then
											RemoveRoleParticipants = False
											blnRunFlag = 1
											Exit Function
										End If 
									End If
								Next								
					
								If oRemoveRoleParticipants.m_Objects("eleEdit").Exist(Environment.Value("SyncTimeOut")) Then
									Call fnInsertResult(Environment.Value("TestCaseID"),"Remove Role Participants Component","Removing Role from Participants","Remove Roles from Participants was successful", "PASS" )
						  		Else
									Call fnInsertResult(Environment.Value("TestCaseID"),"Remove Role Participants Component","Removing Role from Participants","Remove Roles from Participants was unsuccessful", "FAIL")
									RemoveRoleParticipants = False
									blnRunFlag = 1
									Exit Function							 
								End If									
						Next
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_MergeAccount
			
			Private oMergeAccount

			Private Sub Class_Initialize
						Set oMergeAccount = New UI_MergeAccount
			End Sub

			Private Sub Class_Terminate
						Set oMergeAccount = Nothing
			End Sub
			
			Public Function MergeAccount()
						'On Error Resume Nex
						MergeAccount = True
						
						blnFlag = fnSelectActionAccount("AccountActions",objLnkMergeAccountActions,objEleSelectAccountToMergeTitle) 
						If Not blnFlag Then
							MergeAccount = False
							blnRunFlag = 1
							Exit Function
						End If
						
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
										strClass = UCASE(Left(ColName.Name, 3))
										strClass1 = Left(ColName.Name, 3)
										strDTColName = Mid(ColName.Name, 4)
										
										If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												strValue = ColName.Value																													
											
												If ColName.Name = "tblSearchMergeSelect" Then
												        Environment("AccountNo") = strValue
														blnRetVal = fnSelectItemInTable(strValue,oMergeAccount.m_Objects("tblSearchMergeSelect"))
														
												Else
														blnRetVal = SetPageValueDict(oMergeAccount.m_Objects,oMergeAccount,ColName.Name, strValue)
												End if
												If Not blnRetVal Then
													MergeAccount = False
													blnRunFlag = 1
													Exit Function
												End If 
										End If
							Next
							If oMergeAccount.r_Objects("tblPolicyTerms").Exist(Environment.Value("SyncTimeOut")) Then 							    
								Call fnInsertResult(Environment.Value("TestCaseID"),"Merge Account Component","Merging one Account to another Account","Merging "&Environment("AccountNo")& "was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Merge Account Component","Merging one Account to another Account","Merging "&Environment("AccountNo")& "was Unsuccessful", "FAIL")
								MergeAccount = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
			
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_LocationsChangeActiveStatus
			
			Private oLocationsChangeActiveStatus

			Private Sub Class_Initialize
						Set oLocationsChangeActiveStatus = New UI_LocationsChangeActiveStatus
			End Sub

			Private Sub Class_Terminate
						Set oLocationsChangeActiveStatus = Nothing
			End Sub
			
			Public Function LocationsChangeActiveStatus()
						'On Error Resume Nex
						LocationsChangeActiveStatus = True
						
						blnFlag = fnClickQuickLaunchLinks(objEleLocations,objEleAccountFileLocTitle)	
						If Not blnFlag Then
							LocationsChangeActiveStatus = False
							blnRunFlag = 1
							Exit Function
						End If										
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
										strClass = UCASE(Left(ColName.Name, 3))
										strClass1 = Left(ColName.Name, 3)
										strDTColName = Mid(ColName.Name, 4)										
										If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												strValue = ColName.Value																											
												If ColName.Name = "tblLocation" Then
												        Environment("LocationToSelect") = strValue
												        Environment("BeforeEdtActiveStatus") = ""
												        Environment("BeforeEdtActiveStatus") = fnVerifyActiveStatus(oLocationsChangeActiveStatus.m_Objects("tblLocation"),strValue)
														blnRetVal = fnSelectRequiredCheckBox(oLocationsChangeActiveStatus.m_Objects("tblLocation"),strValue)
												Else
														blnRetVal = SetPageValueDict(oLocationsChangeActiveStatus.m_Objects,oLocationsChangeActiveStatus,ColName.Name, strValue)																											
												End if
												If Not blnRetVal Then
													LocationsChangeActiveStatus = False
													blnRunFlag = 1
													Exit Function
												End If 
										End If
							Next
'							Environment("AfterEdtActiveStatus") = ""
'							Environment("AfterEdtActiveStatus") = fnVerifyActiveStatus(oLocationChangeActiveStatus.m_Objects("tblLocation"),Environment("LocationToSelect"))
							If oLocationsChangeActiveStatus.m_Objects("tblLocation").Exist(Environment.Value("SyncTimeOut"))Then	
'							If InStr(Environment("BeforeEdtActiveStatus"),Environment("AfterEdtActiveStatus")) = 0 Then	
								 Call fnInsertResult(Environment.Value("TestCaseID"),"Location Change Active Status Component","Active Status should be changed as Yes/No","Active Status changed successfully", "PASS" )
								'Call fnInsertResult(Environment.Value("TestCaseID"),"Location Change Active Status Component","Active Status should be changed as Yes/No","Before changing Active Status ->"&Environment("BeforeEdtActiveStatus")& "and After changing Active Status"&Environment("AfterEdtActiveStatus"), "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Location Change Active Status Component","Active Status should be changed as Yes/No","Active Status change was unsuccessful", "FAIL")
								'Call fnInsertResult(Environment.Value("TestCaseID"),"Location Change Active Status Component","Active Status should be changed as Yes/No","Before changing Active Status ->"&Environment("BeforeEdtActiveStatus")& "and After changing Active Status"&Environment("AfterEdtActiveStatus"), "FAIL")
								LocationsChangeActiveStatus = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
			
			'Function to verify Active status
			Public Function fnVerifyActiveStatus(objTable,strValue)
				strActiveStatus = ""
				intRows = objTable.RowCount
				If intRows > 0 Then
				   For intI = 1 To intRows
					  If Trim(objTable.GetCellData(intI,4)) = Trim(strValue) Then
						 strActiveStatus = Trim(objTable.GetCellData(intI,2))
						 Exit For
					  End If
				   Next
				End If				
				fnVerifyActiveStatus = strActiveStatus
			End Function
			
			'Function to select the specified checkbox
			Public Function fnSelectRequiredCheckBox(objTable,strValue)
			    blnFlag = False
				intRows = objTable.RowCount
				If intRows > 0 Then
				   For intI = 1 To intRows
					  If Trim(objTable.GetCellData(intI,4)) = Trim(strValue) Then
					     Setting.WebPackage("ReplayType") = 2
					     objTable.ChildItem(intI,1,"Image",0).Click
					     Setting.WebPackage("ReplayType") = 1
					     blnFlag = True
					     Exit For
					  End If
				   Next
				End If
				
				fnSelectRequiredCheckBox = blnFlag
			End Function
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_LocationsSetAsPrimary
			
			Private oLocationsSetAsPrimary

			Private Sub Class_Initialize
						Set oLocationsSetAsPrimary = New UI_LocationsSetAsPrimary
			End Sub

			Private Sub Class_Terminate
						Set oLocationsSetAsPrimary = Nothing
			End Sub
			
			Public Function LocationsSetAsPrimary()
						'On Error Resume Nex
						LocationsSetAsPrimary = True
						
						blnFlag = fnClickQuickLaunchLinks(objEleLocations,objEleAccountFileLocTitle)	
						If Not blnFlag Then
							LocationsSetAsPrimary = False
							blnRunFlag = 1
							Exit Function
						End If										
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
										strClass = UCASE(Left(ColName.Name, 3))
										strClass1 = Left(ColName.Name, 3)
										strDTColName = Mid(ColName.Name, 4)										
										If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
												strValue = ColName.Value																											
												If ColName.Name = "tblLocation" Then												       
														blnRetVal = fnSelectRequiredCheckBox(oLocationsSetAsPrimary.m_Objects("tblLocation"),strValue)
												Else
														blnRetVal = SetPageValueDict(oLocationsSetAsPrimary.m_Objects,oLocationsSetAsPrimary,ColName.Name, strValue)																											
												End if
												If Not blnRetVal Then
													LocationsSetAsPrimary = False
													blnRunFlag = 1
													Exit Function
												End If 
										End If
							Next
							If oLocationsSetAsPrimary.m_Objects("tblLocation").Exist(Environment.Value("SyncTimeOut"))Then	
								 Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Set as Primary Component","Primary Location should be changed","Setting location as primary was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Set as Primary Component","Primary Location should be changed","Setting location as primary was unsuccessful", "FAIL")
								LocationsSetAsPrimary = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
			
			'Function to verify Primary
			Public Function fnVerifyPrimary(objTable,strValue)
				blnFlag = False
				intRows = objTable.RowCount
				If intRows > 0 Then
				   For intI = 1 To intRows					  
					  If Trim(objTable.GetCellData(intI,4)) = Trim(strValue) Then					  	
						 blnFlag = True
						 Exit For
					  End If
				   Next
				End If				
				fnVerifyPrimary = blnFlag
			End Function
			
			'Function to select the specified checkbox
			Public Function fnSelectRequiredCheckBox(objTable,strValue)
			    blnFlag = False
				intRows = objTable.RowCount
				If intRows > 0 Then
				   For intI = 1 To intRows
					  If Trim(objTable.GetCellData(intI,4)) = Trim(strValue) Then
					    If Trim(objTable.GetCellData(intI,2)) = "Yes" Then
						     Setting.WebPackage("ReplayType") = 2
						     objTable.ChildItem(intI,1,"Image",0).Click
						     Setting.WebPackage("ReplayType") = 1
						     blnFlag = True
						     Exit For						     
					    End If 
					  End If
				   Next
				End If
				
				fnSelectRequiredCheckBox = blnFlag
			End Function
End Class

'==================================================================================================
'
'
'==================================================================================================

Class SCR_CompleteActivity
			
			Private oCompleteActivity

			Private Sub Class_Initialize
						Set oCompleteActivity = New UI_CompleteActivity
			End Sub

			Private Sub Class_Terminate
						Set oCompleteActivity = Nothing
			End Sub
			
			Public Function CompleteActivity()
						'On Error Resume Nex
						blnFlag = False
						CompleteActivity=True
						blnFlag = fnMenuNavigation("DeskTop")
						If Not blnFlag Then
						   SearchAccount = False
						   blnRunFlag = 1
						   Exit Function
						End If		
						
						Setting.WebPackage("ReplayType") = 2
						blnFlag = fnClickQuickLaunchLinks(oCompleteActivity.m_Objects("eleMyActivities"),oCompleteActivity.m_Objects("eleActivitySummary"))
						Setting.WebPackage("ReplayType") = 1
						
						If Not blnFlag Then
						   SearchAccount = False
						   blnRunFlag = 1
						   Exit Function
						End If

						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
										strClass = UCASE(Left(ColName.Name, 3))
										strClass1 = Left(ColName.Name, 3)
										strDTColName = Mid(ColName.Name, 4)
										
										If strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" or strClass1 = "tbl" Then
												strValue = ColName.Value																													
												If ColName.Name = "tblSelectActivity" Then
													Wait(5)
													  blnFlag = fnTableSelectSpecifiedCheckBox(objActivitiesTable,strValue)
																						
												ElseIf ColName.Name = "tblVerifyActivity" Then
													Wait(5)
														strCol = 8
													  blnFlag = fnReadTableData(objActivitiesTable,strCol,strValue)
																							
												Else
												
												blnRetVal = SetPageValueDict(oCompleteActivity.m_Objects,oCompleteActivity,ColName.Name,strValue)
												
												End if		
												
												If Not blnFlag Then
													SearchAccount = False
													blnRunFlag = 1
													Exit Function
												End If 
										End If
							Next
							If blnRunFlag <> 1 Then
								Call fnInsertResult(Environment.Value("TestCaseID"),"Complete Activity Component","Complete the Open Activity in Desktop"," Complete the Open Activity was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Complete Activity Component","Complete the Open Activity in Desktop","Complete the Open Activity was unsuccessful", "FAIL")
							End If									
						Next
			End Function
	
	End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_LocationsTransactionNosNavigate
			
			Private oLocationsTransactionNosNavigate

			Private Sub Class_Initialize
						Set oLocationsTransactionNosNavigate = New UI_LocationsTransactionNosNavigate
			End Sub

			Private Sub Class_Terminate
						Set oLocationsTransactionNosNavigate= Nothing
			End Sub
			
			Public Function LocationsTransactionNosNavigate()
						'On Error Resume Next
						LocationsTransactionNosNavigate = True
						
						blnFlag = fnClickQuickLaunchLinks(objEleLocations,objEleAccountFileLocTitle)	
						If Not blnFlag Then
							LocationsTransactionNosNavigate = False
							blnRunFlag = 1
							Exit Function
						End If										
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
								strClass = UCASE(Left(ColName.Name, 3))
								strClass1 = Left(ColName.Name, 3)
								strDTColName = Mid(ColName.Name, 4)										
								If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
										strValue = ColName.Value																											
										If ColName.Name = "tblTransaction" Then												       
												blnRetVal = fnTableClickSpecifiedItemNavigate("Transaction",strValue)
										Else
												blnRetVal = SetPageValueDict(oLocationsTransactionNosNavigate.m_Objects,oLocationsTransactionNosNavigate,ColName.Name,strValue)																											
										End if
										If Not blnRetVal Then
											LocationsTransactionNosNavigate = False
											blnRunFlag = 1
											Exit Function
										End If 
								End If
							Next
							If blnRetVal Then	
								 Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Click Transaction Nos and Navigate Component","Click Transaction Number should navigate to respective page based on the status","Navigation to expected page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Click Transaction Nos and Navigate Component","Click Transaction Number should navigate to respective page based on the status","Navigation to expected page was unsuccessful", "FAIL")
								LocationsTransactionNosNavigate = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
		
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_LocationsTransactionTypNavigate
			
			Private oLocationsTransactionTypNavigate

			Private Sub Class_Initialize
						Set oLocationsTransactionTypNavigate = New UI_LocationsTransactionTypNavigate
			End Sub

			Private Sub Class_Terminate
						Set oLocationsTransactionTypNavigate= Nothing
			End Sub
			
			Public Function LocationsTransactionTypNavigate()
						'On Error Resume Next
						LocationsTransactionTypNavigate = True
						
						blnFlag = fnClickQuickLaunchLinks(objEleLocations,objEleAccountFileLocTitle)	
						If Not blnFlag Then
							LocationsTransactionTypNavigate = False
							blnRunFlag = 1
							Exit Function
						End If										
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
								strClass = UCASE(Left(ColName.Name, 3))
								strClass1 = Left(ColName.Name, 3)
								strDTColName = Mid(ColName.Name, 4)										
								If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
										strValue = ColName.Value																											
										If ColName.Name = "tblTransaction" Then												       
												blnRetVal = fnTableClickSpecifiedItemNavigate("TransactionType",strValue)
										Else
												blnRetVal = SetPageValueDict(oLocationsTransactionTypNavigate.m_Objects,oLocationsTransactionTypNavigate,ColName.Name,strValue)																											
										End if
										If Not blnRetVal Then
											LocationsTransactionTypNavigate = False
											blnRunFlag = 1
											Exit Function
										End If 
								End If
							Next
							If blnRetVal Then	
								 Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Click Transaction Type and Navigate Component","Click Transaction Type should navigate to respective page based on the status","Navigation to expected page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Click Transaction Type and Navigate Component","Click Transaction Type should navigate to respective page based on the status","Navigation to expected page was unsuccessful", "FAIL")
								LocationsTransactionTypNavigate = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
		
End Class


'==================================================================================================
'
'
'==================================================================================================
Class SCR_LocationsPolicyNavigate
			
			Private oLocationsPolicyNavigate

			Private Sub Class_Initialize
						Set oLocationsPolicyNavigate = New UI_LocationsPolicyNavigate
			End Sub

			Private Sub Class_Terminate
						Set oLocationsPolicyNavigate= Nothing
			End Sub
			
			Public Function LocationsPolicyNavigate()
						'On Error Resume Next
						LocationsPolicyNavigate = True
						
						blnFlag = fnClickQuickLaunchLinks(objEleLocations,objEleAccountFileLocTitle)	
						If Not blnFlag Then
							LocationsPolicyNavigate = False
							blnRunFlag = 1
							Exit Function
						End If										
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
								strClass = UCASE(Left(ColName.Name, 3))
								strClass1 = Left(ColName.Name, 3)
								strDTColName = Mid(ColName.Name, 4)										
								If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
										strValue = ColName.Value																											
										If ColName.Name = "tblPolicy" Then												       
												blnRetVal = fnTableClickSpecifiedItemNavigate("Policy",strValue)
										Else
												blnRetVal = SetPageValueDict(oLocationsPolicyNavigate.m_Objects,oLocationsPolicyNavigate,ColName.Name,strValue)																											
										End if
										If Not blnRetVal Then
											LocationsPolicyNavigate = False
											blnRunFlag = 1
											Exit Function
										End If 
								End If
							Next
							If blnRetVal Then	
								 Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Click Policy Navigate Component","Click Policy should navigate to respective page based on the status","Navigation to expected page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Locations Click Policy and Navigate Component","Click Policy should navigate to respective page based on the status","Navigation to expected page was unsuccessful", "FAIL")
								LocationsPolicyNavigate = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
		
End Class

'==================================================================================================
'
'
'==================================================================================================
Class SCR_SubmissionMgrSubmissionNavigate
			
			Private oSubmissionMgrSubmissionNavigate

			Private Sub Class_Initialize
						Set oSubmissionMgrSubmissionNavigate = New UI_SubmissionMgrSubmissionNavigate
			End Sub

			Private Sub Class_Terminate
						Set oSubmissionMgrSubmissionNavigate= Nothing
			End Sub
			
			Public Function SubmissionMgrSubmissionNavigate()
						'On Error Resume Next
						SubmissionMgrSubmissionNavigate = True
						
						blnFlag = fnClickQuickLaunchLinks(objEleSubmissionManager,objEleSubmissionManagerTitle)	
						If Not blnFlag Then
							SubmissionMgrSubmissionNavigate = False
							blnRunFlag = 1
							Exit Function
						End If										
						For Each intKey in dictKeys'loop iterates number of test data records
							rsSheet.Filter = "RecordNum = '" & intKey & "' AND ID = '" & strID & "'"									
							For Each ColName in rsSheet.Fields
								strClass = UCASE(Left(ColName.Name, 3))
								strClass1 = Left(ColName.Name, 3)
								strDTColName = Mid(ColName.Name, 4)										
								If strClass1 = "tbl" or strClass1 = "edt" or strClass1 = "ele" or strClass1 = "lst" or strClass1 = "img" or strClass1 = "CHK" or strClass1 = "RGD" or strClass1 = "lnk" or strClass1 = "btn" Then
										strValue = ColName.Value																											
										If ColName.Name = "tblSubmission" Then												       
												blnRetVal = fnClickSubmissionNavigate(objTblSubmissionManager,strValue)
										Else
												blnRetVal = SetPageValueDict(oSubmissionMgrSubmissionNavigate.m_Objects,oSubmissionMgrSubmissionNavigate,ColName.Name,strValue)																											
										End if
										If Not blnRetVal Then
											SubmissionMgrSubmissionNavigate = False
											blnRunFlag = 1
											Exit Function
										End If 
								End If
							Next
							If blnRetVal Then	
								 Call fnInsertResult(Environment.Value("TestCaseID"),"Submission Manager Click Submission and  Navigate Component","Click Submission should navigate to respective page based on the status","Navigation to expected page was successful", "PASS" )
							Else
								Call fnInsertResult(Environment.Value("TestCaseID"),"Submission Manager Click Submission and Navigate Component","Click Submission should navigate to respective page based on the status","Navigation to expected page was unsuccessful", "FAIL")
								SubmissionMgrSubmissionNavigate = False
								blnRunFlag = 1
								Exit Function
							End If									
						Next
			End Function
		
		'Function to click Submission and navigate to its specific page
		Public Function fnClickSubmissionNavigate(objTable,strSelectionValue)
			blnFlag = False
			intRows = objTable.RowCount
			If intRows > 0 Then
			   For intI = 1 to intRows     
				  If Trim(objTable.GetCellData(intI,3)) = Trim(strSelectionValue) Then
				  	 Environment("SubmissionStatus") = objTable.GetCellData(intI,6)
					 Setting.WebPackage("ReplayType") = 2
					 objTable.ChildItem(intI,3,"WebElement",1).Click
					 Setting.WebPackage("ReplayType") = 1		 
					 Exit For		
				  End If
			   Next
			End If
			
			Select Case Environment("SubmissionStatus")
				Case "Draft","Not-taken","Withdrawn","Expired"
				   Set objToVerify = objEleQualification
				Case "Quoted","Bound"
				   Set objToVerify = objEleQuoteTitle
			End Select
			
			'Verify whether page navigation to expected page is successful
			If objToVerify.Exist(Environment("SyncTimeOut")) Then  
			 	blnFlag = True
			End If			
			fnClickSubmissionNavigate = blnFlag
		 End Function
End Class
