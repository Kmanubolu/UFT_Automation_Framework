'==================================================================================================
'
'
'==================================================================================================
Class UI_Login

		Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "edtUserName", m_Objects("pgeGuidewirePolicyCenter").WebEdit("name:=Login:LoginScreen:LoginDV:username")
					.Add "edtPassword", m_Objects("pgeGuidewirePolicyCenter").WebEdit("name:=Login:LoginScreen:LoginDV:password")
					.Add "eleLogin", m_Objects("pgeGuidewirePolicyCenter").WebElement("html id:=Login:LoginScreen:LoginDV:submit-btnInnerEl")
				End With
				
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleTabMenu", r_Objects("pgeGuidewirePolicyCenter").WebElement("html id:=TabBar:DesktopTab")
				End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_Logoff

		Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						'.Add "elePereference", m_Objects("pgeGuidewirePolicyCenter").WebElement("html id:=:TabLinkMenuButton-btnIconEl","index:=0")
						
						.Add "elePereference", m_Objects("pgeGuidewirePolicyCenter").WebElement("class:=x-btn-icon-el g-preferences-icon ","index:=0")
						.Add "eleLogOut", m_Objects("pgeGuidewirePolicyCenter").WebElement("html id:=TabBar:LogoutTabBarLink-textEl","index:=0")
						
						'.Add "eleOk", m_Objects("pgeGuidewirePolicyCenter").WebElement("html id:=button-1005-btnInnerEl")
						.Add "eleOk", Browser("micclass:=Browser").Page("micclass:=Page").WebElement("html id:=button-1005-btnInnerEl")
						
						'.Add "btnOk", m_Objects("pgeGuidewirePolicyCenter").Dialog("text:=Windows Internet Explorer").WinButton("text:=&Leave this page")
						.Add "btnOk", Browser("micclass:=Browser").Dialog("text:=Windows Internet Explorer").WinButton("regexpwndtitle:=&Leave this page")
						
					End With

					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "edtUserName", r_Objects("pgeGuidewirePolicyCenter").WebEdit("name:=Login:LoginScreen:LoginDV:username")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class
'==================================================================================================
'
'
'==================================================================================================
Class UI_NewAccountSearch

		Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "edtCompanyName", m_objects("pgeGuidewirePolicyCenter").WebEdit("name:=NewAccount:NewAccountScreen:NewAccountSearchDV:GlobalContactNameInputSet:Name")
						.Add "lnkSearch", m_objects("pgeGuidewirePolicyCenter").Link("name:=Search")
						.Add "bliCreateNewAccount", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=NewAccount:NewAccountScreen:NewAccountButton-btnInnerEl")						
						.Add "lnkCompany", m_objects("pgeGuidewirePolicyCenter").Link("html id:=NewAccount:NewAccountScreen:NewAccountButton:NewAccount_Company-textEl")
						.Add "lnkPerson", m_objects("pgeGuidewirePolicyCenter").Link("html id:=NewAccount:NewAccountScreen:NewAccountButton:NewAccount_Person-itemEl")
					End With

					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleCreateAccount", r_Objects("pgeGuidewirePolicyCenter").WebElement("html id:=CreateAccount:CreateAccountScreen:ttlBar")
						.Add "edtCompanyName", r_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalContactNameInputSet:Name-inputEl","name:=.*GlobalContactNameInputSet:Name","index:=0")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_CreateAccount

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "edtFirstName", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*FirstName-inputEl")
						.Add "edtLastName", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*LastName-inputEl")
						.Add "edtAddressLine1", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*AddressLine1-inputEl","index:=0")
						.Add "edtCity", m_objects("pgeGuidewirePolicyCenter").WebEdit("name:=.*City")
						.Add "lstState", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*State-inputEl")
						.Add "eleCreateAccount1", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*CreateAccountScreen:ttlBar")
						.Add "edtPostalCode", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*PostalCode-inputEl","index:=0")
						.Add "lstAddressType", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*CreateAccountDV:AddressType-inputEl","name:=.*CreateAccountDV:AddressType","index:=0")
						.Add "eleOrganization", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*ProducerSelectionInputSet:Producer:SelectOrganization")
						.Add "edtOrganizationName", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*Name-inputEl")
						.Add "lnkSearchOrg", m_objects("pgeGuidewirePolicyCenter").Link("html id:=.*SearchAndResetInputSet:SearchLinksInputSet:Search")
						.Add "eleSelect", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*OrganizationSearchResultsLV:0:_Select")
					    .Add "lstProducerCode", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*ProducerSelectionInputSet:ProducerCode-inputEl","index:=0")
                       	.Add "eleUpdate", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*CreateAccountScreen:Update-btnInnerEl")
						.Add "eleCancel", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=CreateAccount:CreateAccountScreen:Cancel")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAccountNumber", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile:AccountFileInfoBar:Account-btnInnerEl")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_SearchAccount

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleSearchAccounts",m_objects("pgeGuidewirePolicyCenter").WebElement("class:=x-tree-node-text","innertext:=Accounts","index:=1")
					.Add "edtSearchEnterAccountNo", m_objects("pgeGuidewirePolicyCenter").WebEdit("name:=AccountSearch:AccountSearchScreen:AccountSearchDV:AccountNumber")					
				End With
				
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleAccountSummaryAccountNo",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Summary:AccountFile_SummaryScreen:AccountFile_Summary_BasicInfoDV:AccountNumber-inputEl.*")
				End With
			
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_NewSubmissions

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "edtEffectiveDate",m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*DefaultPPEffDate-inputEl","name:=.*ProductSettingsDV:DefaultPPEffDate")						
				End With

				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleAccountSummaryAccountNo",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Summary:AccountFile_SummaryScreen:AccountFile_Summary_BasicInfoDV:AccountNumber-inputEl.*")
				End With	
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class



'==================================================================================================
'
'
'==================================================================================================
Class UI_Qualifications

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleNext", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Next-btnInnerEl")					
				End With

				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "elePolicyInfo",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:LOBWizardStepGroup:SubmissionWizard_PolicyInfoScreen:ttlBar")
				End With			
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_PolicyInfo

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "edtFEIN", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=SubmissionWizard:LOBWizardStepGroup:SubmissionWizard_PolicyInfoScreen:SubmissionWizard_PolicyInfoDV:AccountInfoInputSet:PolicyOfficialIDInputSet:OfficialIDDV_Input-inputEl")
					.Add "edtSSN", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=SubmissionWizard:LOBWizardStepGroup:SubmissionWizard_PolicyInfoScreen:SubmissionWizard_PolicyInfoDV:AccountInfoInputSet:PolicyOfficialIDInputSet:OfficialIDDV_Input-inputEl")
					.Add "edtIndustryCode", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=SubmissionWizard:LOBWizardStepGroup:SubmissionWizard_PolicyInfoScreen:SubmissionWizard_PolicyInfoDV:AccountInfoInputSet:IndustryCode-inputEl")
					.Add "lstOrganizationType", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=SubmissionWizard:LOBWizardStepGroup:SubmissionWizard_PolicyInfoScreen:SubmissionWizard_PolicyInfoDV:AccountInfoInputSet:OrganizationType-inputEl")
					.Add "eleNext", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Next-btnInnerEl")
				End With
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleLocations",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:LOBWizardStepGroup:LineWizardStepSet:LocationsScreen:ttlBar")
				End With	
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_Locations

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleNext", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Next-btnInnerEl")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleWCCoverage", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:LOBWizardStepGroup:LineWizardStepSet:WorkersCompCoverageConfigScreen:ttlBar")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_WCCoverage

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAddClass", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:LOBWizardStepGroup:LineWizardStepSet:WorkersCompCoverageConfigScreen:WorkersCompCoverageCV:WorkersCompClassesInputSet:WCCovEmpLV_tb:Add-btnInnerEl")
						.Add "tblAddClass", m_objects("pgeGuidewirePolicyCenter").WebTable("Cols:=8")
						.Add "eleNext", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Next-btnInnerEl")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleSupplimental", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:LOBWizardStepGroup:LineWizardStepSet:WorkersCompSupplementalScreen:ttlBar")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_Supplemental

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleNext", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Next-btnInnerEl")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleWCOptions", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=gridcolumn.*","innertext:=WC Options","index:=0")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class


'==================================================================================================
'
'
'==================================================================================================
Class UI_WCOptions

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleNext", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Next-btnInnerEl")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleRiskAnalysis", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Job_RiskAnalysisScreen:0")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class


'==================================================================================================
'
'
'==================================================================================================
Class UI_RiskAnalysis

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleNext", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:Next-btnInnerEl")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "elePolicyReview", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:SubmissionWizard_PolicyReviewScreen:ttlBar")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class



'==================================================================================================
'
'
'==================================================================================================
Class UI_PolicyReview

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleQuote", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:SubmissionWizard_PolicyReviewScreen:JobWizardToolbarButtonSet:QuoteOrReview-btnInnerEl")
						.Add "eleQuoteAlert", m_objects("pgeGuidewirePolicyCenter").WebElement("class:=x-container x-container-default x-table-layout-ct","innertext:=This quote will require underwriting approval prior to binding\.")
						.Add "eleRiskAnalysisTree", m_objects("pgeGuidewirePolicyCenter").WebElement("class:=x-tree-node-text ","innertext:=Risk Analysis")
						.Add "tblRiskAnalysisIssuesTable", m_objects("pgeGuidewirePolicyCenter").WebTable("class:=x-gridview-.*-table x-grid-table x-grid-with-col-lines x-grid-with-row-lines","cols:=5")
						.Add "eleRiskAnalysisApprove", m_objects("pgeGuidewirePolicyCenter").WebElement("class:=x-btn-inner x-btn-inner-center", "innertext:=Approve")
						.Add "eleRiskAnalysisOK", m_objects("pgeGuidewirePolicyCenter").WebElement("class:=x-btn-inner x-btn-inner-center","innertext:=OK")
						.Add "eleTreeQuote", m_objects("pgeGuidewirePolicyCenter").WebElement("class:=x-tree-node-text ","innertext:=Quote")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleQuote", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:SubmissionWizard_QuoteScreen:ttlBar")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class



'==================================================================================================
'
'
'==================================================================================================
Class UI_Quote

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "bliBindOptions", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SubmissionWizard:SubmissionWizard_QuoteScreen:JobWizardToolbarButtonSet:BindOptions-btnInnerEl")
						.Add "eleAreYouSureOK", m_objects("pgeGuidewirePolicyCenter").WebElement("class:=x-btn-inner x-btn-inner-center","innertext:=OK")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleSubmissionBound", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=JobComplete:JobCompleteScreen:ttlBar")
						.Add "tblSubmissionInfo", r_objects("pgeGuidewirePolicyCenter").WebTable("column names:=Submission Bound","index:=0")

					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class


'==================================================================================================
'
'
'==================================================================================================
Class UI_UpdateAccount

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleEditAccount", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Summary:AccountFile_SummaryScreen:EditAccount-btnInnerEl","innertext:=Edit Account")
						.Add "eleEditAccountTitle", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=EditAccountPopup:EditAccountScreen:ttlBar","innertext:=Edit Account")
						.Add "edtFirstName", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalPersonNameInputSet:FirstName-inputEl","name:=.*GlobalPersonNameInputSet:FirstName")
						.Add "edtLastName", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalPersonNameInputSet:LastName-inputEl","name:=.*GlobalPersonNameInputSet:LastName")						
						.Add "edtCompanyName", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalContactNameInputSet:Name-inputEl","name:=.*GlobalContactNameInputSet:Name")											
						.Add "edtAddressLine1", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalAddressInputSet:AddressLine1-inputEl","name:=.*GlobalAddressInputSet:AddressLine1","index:=0")
						.Add "edtCity", m_objects("pgeGuidewirePolicyCenter").WebEdit("name:=.*GlobalAddressInputSet:City","html id:=.*GlobalAddressInputSet:City-inputEl")
						.Add "lstState", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalAddressInputSet:State-inputEl","name:=.*GlobalAddressInputSet:State")
						.Add "edtPostalCode", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalAddressInputSet:PostalCode-inputEl","index:=0")
						.Add "lstAddressType", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*AddressType-inputEl","name:=.*AddressType","index:=0")
						.Add "eleUpdate", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=EditAccountPopup:EditAccountScreen:Update")
						.Add "eleCancel", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=EditAccountPopup:EditAccountScreen:Cancel")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAccountSummary", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Summary:AccountFile_SummaryScreen:ttlBar","innertext:=Account File Summary")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class
'==================================================================================================
'
'
'==================================================================================================
Class UI_AddRelatedAccounts

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAddBtn",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_RelatedAccounts:AddRelatedAccount-btnInnerEl","innertext:=Add","index:=0")
						.Add "eleAccountRelationship",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=RelatedAccountPopup:ttlBar","innertext:=Account Relationship","index:=0")
						.Add "lstRelationship", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=RelatedAccountPopup:RelationshipType-inputEl","name:=RelatedAccountPopup:RelationshipType")
						.Add "edtRelatedAccount", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=RelatedAccountPopup:RelatedAccount-inputEl","name:=RelatedAccountPopup:RelatedAccount")
						.Add "eleName", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=RelatedAccountPopup:Name-labelEl","innertext:=Name")											
						.Add "eleUpdate", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=RelatedAccountPopup:Update-btnInnerEl")
						.Add "eleCancel", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=RelatedAccountPopup:Cancel-btnInnerEl")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAccountFileRelatedAccounts", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_RelatedAccounts:ttlBar","innertext:=Account File Related Accounts")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class


'==================================================================================================
'
'
'==================================================================================================
Class UI_RemoveRelatedAccounts

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleRemoveBtn",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_RelatedAccounts:RemoveRelatedAccount-btnInnerEl","innertext:=Remove","index:=0")						
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAccountFileRelatedAccounts", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_RelatedAccounts:ttlBar","innertext:=Account File Related Accounts")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class



'==================================================================================================
'
'
'==================================================================================================
Class UI_CreateLocation

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleLocation", m_objects("pgeGuidewirePolicyCenter").WebElement("innerhtml:=Locations","html tag:=SPAN")
						.Add "eleAddNewLocation", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Locations:AccountFile_LocationsScreen:AccountFile_LocationsLV_tb:AddNewLocation-btnInnerEl")
						.Add "eleLocationInformation", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountLocationPopup:LocationScreen:ttlBar")
						
						.Add "btnNonSpecificLocationYES", m_objects("pgeGuidewirePolicyCenter").WebButton("html id:=AccountLocationPopup:LocationScreen:AccountLocationDetailInputSet:NonSpecificLocation_true-inputEl")										
						.Add "btnNonSpecificLocationNO", m_objects("pgeGuidewirePolicyCenter").WebButton("html id:=AccountLocationPopup:LocationScreen:AccountLocationDetailInputSet:NonSpecificLocation_false-inputEl")
						.Add "edtAddressLine1", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=AccountLocationPopup:LocationScreen:AccountLocationDetailInputSet:AddressInputSet:globalAddressContainer:GlobalAddressInputSet:AddressLine1-inputEl")
						.Add "edtCity", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=AccountLocationPopup:LocationScreen:AccountLocationDetailInputSet:AddressInputSet:globalAddressContainer:GlobalAddressInputSet:City-inputEl")
						.Add "lstState", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=AccountLocationPopup:LocationScreen:AccountLocationDetailInputSet:AddressInputSet:globalAddressContainer:GlobalAddressInputSet:State-inputEl")
						.Add "edtPostalCode", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=AccountLocationPopup:LocationScreen:AccountLocationDetailInputSet:AddressInputSet:globalAddressContainer:GlobalAddressInputSet:PostalCode-inputEl")
						.Add "eleUpdate", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountLocationPopup:LocationScreen:Update-btnInnerEl")
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAccountFileLocations", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Locations:AccountFile_LocationsScreen:ttlBar")
						.Add "tblLocationsList", r_objects("pgeGuidewirePolicyCenter").WebTable("cols:=8")		

					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_DeleteLocation

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleLocation", m_objects("pgeGuidewirePolicyCenter").WebElement("innerhtml:=Locations","html tag:=SPAN")
						.Add "eleRemoveLocation", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Locations:AccountFile_LocationsScreen:AccountFile_LocationsLV_tb:removeLocation-btnInnerEl")
						
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAccountFileLocations", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Locations:AccountFile_LocationsScreen:ttlBar")
						.Add "tblLocationsList", r_objects("pgeGuidewirePolicyCenter").WebTable("cols:=8")		

					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class


'==================================================================================================
'
'
'==================================================================================================

Class UI_SearchAddressBook

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "lstType",m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*ContactSearchScreen:ContactType-inputEl","name:=.*ContactSearchScreen:ContactType","index:=0")
					.Add "edtCompanyName",m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalContactNameInputSet:Name-inputEl","name:=.*GlobalContactNameInputSet:Name","index:=0")
					.Add "edtFirstName",m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalPersonNameInputSet:FirstName-inputEl","name:=.*GlobalPersonNameInputSet:FirstName","index:=0")
					.Add "edtLastName",m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalPersonNameInputSet:LastName-inputEl","name:=.*GlobalPersonNameInputSet:LastName","index:=0")
					.Add "lstState",m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*GlobalAddressInputSet:State-inputEl","name:=.*GlobalAddressInputSet:State","index:=0")
					.Add "lnkSearch",m_objects("pgeGuidewirePolicyCenter").Link("html id:=.*SearchLinksInputSet:Search","name:=Search")
					.Add "tblAddressBookSelect",m_objects("pgeGuidewirePolicyCenter").WebTable("html id:=gridview.*","html tag:=TABLE")					
				End With
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleCreateAccountTitle",r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*CreateAccountScreen:ttlBar","innertext:=Create account")
				End With	
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class


'==================================================================================================
'
'
'==================================================================================================
Class UI_SearchRelatedAccount

Public m_Objects,r_Objects

		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
					Set m_Objects = CreateObject("Scripting.Dictionary")
					Set r_Objects = CreateObject("Scripting.Dictionary")
					'do action on application
					With m_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAddBtn",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_RelatedAccounts:AddRelatedAccount-btnInnerEl","innertext:=Add","index:=0")
						.Add "eleRelatedAccountSearch", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=RelatedAccountPopup:RelatedAccount:SelectRelatedAccount","index:=0")
						'.Add "eleSearchRelatedAccountTitle", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=SearchRelatedAccountPopup:ttlBar","innertext:=Search Related Account")
						.Add "edtSearchRelatedAccountAccountNo", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=.*AccountNumber-inputEl","name:=.*AccountSearchDV:AccountNumber")
						.Add "lnkRelatedAccountSearch",m_objects("pgeGuidewirePolicyCenter").link("html id:=.*SearchAndResetInputSet:SearchLinksInputSet:Search","name:=Search")
						.Add "tblRelatedAccountSelect",m_objects("pgeGuidewirePolicyCenter").WebTable("html id:=gridview.*","name:=WebTable","innertext:=Select.*")						
					End With

					'Read Data from App
					With r_Objects
						.Add "pgeGuidewirePolicyCenter", BrowserPage()
						.Add "eleAccountRelationship",m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=RelatedAccountPopup:ttlBar","innertext:=Account Relationship")
					End With
		End Sub

		Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property

End Class


'==================================================================================================
'
'
'==================================================================================================
Class UI_SearchAccountPolicyMove

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "lnkSearchAccounts",m_objects("pgeGuidewirePolicyCenter").Link("html id:=AccountFile_AccountSearch:OtherAccountSearchScreen:AccountSearchDV:SearchAndResetInputSet:SearchLinksInputSet:Search")
					.Add "edtSearchEnterAccountNo", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=AccountFile_AccountSearch:OtherAccountSearchScreen:AccountSearchDV:AccountNumber-inputEl")					
					.Add "tblSelectAccount",m_objects("pgeGuidewirePolicyCenter").WebTable("cols:=4")
				End With
				
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleMovepoliciesSelection",r_objects("pgeGuidewirePolicyCenter").Webelement("html id:=AccountFile_MovePoliciesSelection:ttlBar")
				End With
			
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================
Class UI_AccountPolicyMove

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "tblPolictSelection",m_objects("pgeGuidewirePolicyCenter").WebTable("cols:=7")
					.Add "eleMovePolicyToAccount",m_objects("pgeGuidewirePolicyCenter").Webelement("html id:=AccountFile_MovePoliciesSelection:MovePoliciesButton-btnInnerEl")
					
				End With
				
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleAccountFileSummary",r_objects("pgeGuidewirePolicyCenter").Webelement("html id:=AccountFile_Summary:AccountFile_SummaryScreen:ttlBar")
				End With
			
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class'==================================================================================================
'
'
'==================================================================================================
Class UI_RemoveRoleParticipants

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
							
					'Below One used for Remove Participant
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleEdit", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Edit","html id:=AccountFile_Roles:AccountFile_RolesScreen:Edit-btnInnerEl","index:=0")
					.Add "eleRemove",m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Remove","html id:=AccountFile_Roles:AccountFile_RolesScreen:Remove-btnInnerEl","index:=0")
					.Add "eleUpdate", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Update","html id:=AccountFile_Roles:AccountFile_RolesScreen:Update","index:=0")
			
				End With				
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================

Class UI_AddRoleParticipants

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleEdit", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Edit","html id:=AccountFile_Roles:AccountFile_RolesScreen:Edit-btnInnerEl","index:=0")
					.Add "eleAdd", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Add","html id:=AccountFile_Roles:AccountFile_RolesScreen:Add-btnInnerEl","index:=0")
					.Add "lstAddRole", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=<none>","class:=x-grid-cell-inner","index:=0")
'					.Add "eleAddRolevalue", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Audit Examiner","class:=x-boundlist-item","index:=0")
					.Add "eleSelectUser", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Select User...","html id:=AccountFile_Roles:AccountFile_RolesScreen:AccountRolesLV:1:AssignedUser:UserBrowseMenuItem","index:=0")
					.Add "edtUserName", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=UserSearchPopup:UserSearchPopupScreen:UserSearchDV:Username-inputEl","index:=0")
					.Add "lnkSearch", m_objects("pgeGuidewirePolicyCenter").Link("innertext:=Search","html id:=UserSearchPopup:UserSearchPopupScreen:UserSearchDV:SearchAndResetInputSet:SearchLinksInputSet:Search","index:=0")
					.Add "eleSelect", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Select","html id:=UserSearchPopup:UserSearchPopupScreen:UserSearchResultsLV:0:_Select","index:=0")
'					.Add "eleAssigneduser", m_objects("pgeGuidewirePolicyCenter").WebElement("html tag:=DIV","class:=g-helper-cell-text","index:=0")
'					.Add "edtAssigneduservalue", m_objects("pgeGuidewirePolicyCenter").WebEdit("html tag:=Input","class:=x-form-field x-form-text x-form-focus x-field-form-focus x-field-default-form-focus","name:=AssignedUser","index:=0")
'					.Add "eleAssignGroupCol", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=gridcolumn.*","innertext:=Assigned Group","index:=0")
					.Add "lstAssignedgroup", m_objects("pgeGuidewirePolicyCenter").WebElement("html tag:=DIV","innertext:=<none>","class:=x-grid-cell-inner")
					.Add "eleUpdate", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=Update","html id:=AccountFile_Roles:AccountFile_RolesScreen:Update-btnInnerEl","index:=0")
					
				End With
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
'					.Add "eleParticipantssummary", r_objects("pgeGuidewirePolicyCenter").WebElement("html id:=AccountFile_Roles:AccountFile_RolesScreen:ttlBar","innertext:=Account File Participants")
				End With	
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class


'==================================================================================================
'
'
'==================================================================================================

Class UI_MergeAccount

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()					
					.Add "edtAccountNo", m_objects("pgeGuidewirePolicyCenter").WebEdit("name:=.*AccountSearchDV:AccountNumber","html id:=.*AccountSearchDV:AccountNumber-inputEl")
					.Add "lnkSearch", m_objects("pgeGuidewirePolicyCenter").Link("name:=Search","html id:=.*SearchAndResetInputSet:SearchLinksInputSet:Search")
					.Add "tblSearchMergeSelect", m_objects("pgeGuidewirePolicyCenter").WebTable("html id:=gridview.*","html tag:=TABLE","name:=WebTable")
					.Add "eleMergeAccounts", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*MergeAccounts-btnInnerEl","innertext:=Merge Accounts")
					.Add "eleOk", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=button-1005-btnInnerEl","innertext:=Ok")
					.Add "eleCancel", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*CancelButton-btnInnerEl","innertext:=Cancel")					
				End With
				With r_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "tblPolicyTerms", r_objects("pgeGuidewirePolicyCenter").WebTable("html id:=gridview.*","html tag:=TABLE","name:=WebTable","index:=2")
				End With	
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class


'==================================================================================================
'
'
'==================================================================================================

Class UI_LocationsChangeActiveStatus

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()	
					.Add "tblLocation", m_objects("pgeGuidewirePolicyCenter").WebTable("html id:=gridview.*","class:=.*x-grid-table x-grid-with-col-lines x-grid-with-row-lines.*","index:=0")					
					.Add "eleChangeActiveStatus", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*AccountFile_LocationsLV_tb:changeActiveStatus-btnInnerEl","innertext:=Change Active Status")					
					.Add "eleLocationCode", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*AccountLocationDetailInputSet:LocationCode-labelEl","innertext:=Location Code")
				End With				
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================

Class UI_LocationsSetAsPrimary

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()	
					.Add "tblLocation", m_objects("pgeGuidewirePolicyCenter").WebTable("html id:=gridview.*","class:=.*x-grid-table x-grid-with-col-lines x-grid-with-row-lines.*","index:=0")					
					.Add "eleSetAsPrimary", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=.*AccountFile_LocationsLV_tb:setToPrimary-btnInnerEl","innertext:=Set As Primary")					
					
				End With				
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================

Class UI_CompleteActivity

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()
					.Add "eleMyActivities", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=My Activities","html tag:=SPAN","class:=x-tree-node-text")
					.Add "eleActivitySummary", m_objects("pgeGuidewirePolicyCenter").WebElement("innertext:=My Activities","html id:=DesktopActivities:DesktopActivitiesScreen:0")
					.Add "lstSelectActivitystatus", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=DesktopActivities:DesktopActivitiesScreen:DesktopActivitiesLV:activitiesFilter-inputEl")
					.Add "eleCompleteButton", m_objects("pgeGuidewirePolicyCenter").WebElement("html id:=DesktopActivities:DesktopActivitiesScreen:DesktopActivitiesLV_tb:DesktopActivities_CompleteButton-btnInnerEl")
					.Add "lstSelectCompleteActivitystatus", m_objects("pgeGuidewirePolicyCenter").WebEdit("html id:=DesktopActivities:DesktopActivitiesScreen:DesktopActivitiesLV:activitiesFilter-inputEl")
				End With
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================

Class UI_LocationsTransactionNosNavigate

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()	
					.Add "tblTransaction", m_objects("pgeGuidewirePolicyCenter").WebTable("class:=.*x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","name:=WebTable","index:=1")									
				End With				
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================

Class UI_LocationsTransactionTypNavigate

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()	
					.Add "tblTransaction", m_objects("pgeGuidewirePolicyCenter").WebTable("class:=.*x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","name:=WebTable","index:=1")									
				End With				
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================

Class UI_LocationsPolicyNavigate

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()	
					.Add "tblPolicy", m_objects("pgeGuidewirePolicyCenter").WebTable("class:=.*x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","name:=WebTable","index:=0")									
				End With				
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

'==================================================================================================
'
'
'==================================================================================================

Class UI_SubmissionMgrSubmissionNavigate

Public m_Objects,r_Objects
		Private Sub Class_Terminate()
				Set m_Objects = Nothing
				Set r_Objects = Nothing
		End Sub

		Private Sub Class_Initialize()
				Set m_Objects = CreateObject("Scripting.Dictionary")
				Set r_Objects = CreateObject("Scripting.Dictionary")
				With m_Objects
					.Add "pgeGuidewirePolicyCenter", BrowserPage()	
					.Add "tblSubmission", m_objects("pgeGuidewirePolicyCenter").WebTable("class:=.*x-grid-with-col-lines x-grid-with-row-lines","html id:=gridview.*","name:=infobar_WC","index:=0")									
				End With				
			End Sub
			Public Property Get Object(oDict, vkey)
				If oDict.Exists(vkey) Then
					oDict(vkey).Init()
					Set Object = oDict.Item(vkey)
				Else
					Set Object = Nothing
				End If
		End Property
	
End Class

