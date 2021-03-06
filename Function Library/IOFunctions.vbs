Public dictKeys, dictComp, rsSheet, rsScenarios, intNum1, blnRunFlag
Public strFrameworkFolderpath,strErrorText,strID,strHwnd,strColumnValue

Set dictComp = CreateObject("Scripting.Dictionary")
Set dictKeys = CreateObject("Scripting.Dictionary")
Environment.Value("AccountNumber1") = ""
Environment.Value("AccountNumber2") = ""
Environment.Value("AccountNumber3") = ""
Environment.Value("AccountNumber4") = ""
Environment.Value("AccountNumber5") = ""
Environment.Value("Module") = ""
Environment("ErrorMsg") = ""
'Set Ado Constants
Const adStateClosed = 0
Const adStateOpen = 1
Const adOpenStatic = 3
Const adUseClient = 3
Const adLockBatchOptimistic = 4


' overridden exit method for all objects
RegisterUserFunc "WebEdit","Exist","ObjExist"
RegisterUserFunc "WebElement","Exist","ObjExist"
RegisterUserFunc "Image","Exist","ObjExist"
RegisterUserFunc "Link","Exist","ObjExist"
RegisterUserFunc "WebButton","Exist","ObjExist"
RegisterUserFunc "WebTable","Exist","ObjExist"

'==================================================================================================
'
'
'==================================================================================================
Public Function ObjExist(obj,intTime)
	ObjExist = False
	If isNumeric(intTime)=False Then  'this condition occurs when the exist method has not parameter passed
		intTime = Environment.Value("ObjectTimeOut")
	End If
	For Iterator = 1 To intTime
		If obj.Exist(1) Then
				ObjExist = True
				Exit Function
		End if
	Next
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function CallMainScript()

			Call fnSummaryInitialization("Policy Center Execution Summary Report")
			'Set dictComp = GetComponentsDictionary(Environment.Value("DataFilePath"),"Select * from [Components$]")
			Set rsScenarios = GetContentsFromDB(Environment.Value("DataFilePath"),"Select * from [TestCases$]")
			rsScenarios.Filter = "ExecutionRequired = 'YES'"
			While rsScenarios.EOF = False
						ExecuteFile (Environment.Value("FrameworkFolderPath") & "Object Dictionary\OBJ_Common.vbs")
						strTestCaseID = GetRecordSetValue(rsScenarios,"TestCaseName")
						Environment("TestCaseID")  = strTestCaseID
						strExcution = GetRecordSetValue(rsScenarios,"ExecutionRequired")
						Environment.Value("TestID") = GetRecordSetValue(rsScenarios,"TestCaseID")
						strID = GetRecordSetValue(rsScenarios,"ID")
						Environment.Value("ID") = strID
						Environment.Value("TestSetId") = GetRecordSetValue(rsScenarios,"TestSetId")
						Call fnInitilization(strModuleName & " Module " & strTestCaseID)
						blnRunFlag = 0
						print "Execution:::" & strTestCaseID
									For Each ColName in rsScenarios.Fields
											strSearchString = "Login,Logoff"
											strColumnValue = Trim(ColName.Value)
											If Instr(ColName.Name,"Com") > 0 and Len(strColumnValue) > 0 Then
												intpos =  Instr(1,strSearchString,strColumnValue,1)
												If blnRunFlag = 0 Then
															If intpos > 0 Then
																		Select case strColumnValue
																		Case "Login"
																					Call CloseOpenBrowser()
																					Call LaunchApp()
																					sClassName="SCR_Login"
																					Execute "Set scInstance=New "& sClassName
																					scInstance.Login Environment.Value("LoginUserName"),Environment.Value("LoginPassword")
																					
																		Case "Logoff"
																					sClassName="SCR_Logoff"
																					Execute"Set scrInstance=New " & sClassName
																					Execute "scrInstance." & strColumnValue
																		End Select
																		If Err.Description<>""or Environment("ErrorMsg")<>""or blnRunFlag<>0Then
																					Call fnInsertResult(strTestCaseID,strColumnValue,Chr(34) & strColumnValue & Chr(34) & " Componrnt should be executed successfully",Chr(34) & strColumnValue & Chr(34) & " Component did not execute successfully","FAIL")
																					'Execute"Set scrInstance = New SCR_Logoff"
																					'scrInstance.Logoff
																					Call CloseOpenBrowser()
																					print "Error:::" & strTestCAseID & " Exiting Due to Error in Component---" & strColumnValue
																					Exit For 'Since Comp has beeb failed
																		End If
															Else
																		Set rsSheet = GetContentsFromDB(Environment.Value("DataFilePath"),"Select*from [" & strColumnValue & "$" & "]")
																		rsSheet.Filter = "ID = '" & strID &"'"
																		For inKey = 0 to rsSheet.RecordCount-1
																					strActualKey = GetRecordSetValue(rsSheet,"RecordNum")
																					If not dictKeys.Exists(strActualKey)Then
																							dictKeys.Add strActualKey,strActualKey
																					End If
																		rsSheet.MoveNext
																		Next
																		If rsSheet.RecordCount > 0Then
																					rsSheet.MoveFirst
																					Call fnInsertResult(Environment.Value("TestCaseID"), strColumnValue & " Component",Chr(34) & strColumnValue & Chr(34) & " Component execution should be started",Chr(34) & strColumnValue & Chr(34) & " Component execution has been started","PASS")
																		Else
																					Call fnInsertResult(Environment.Value("TestCaseID"), strColumnValue & " Component",Chr(34) & strColumnValue & Chr(34) & " Component execution should be started",Chr(34) & strColumnValue & Chr(34) & " Component execution has not started.Verify if testdata has beeen specified in " & Chr(34) & strColumnValue & Chr(34) & "dara sheet","FAIL")
																					Exit For
																		End If
																		sClassName = "SCR_" & strColumnValue
																		Execute "Set scrInstance = New " & sClassName 'Objects are initialized in dictionary
																		Execute "scrInstance." & strColumnValue
																		If Err.Description <> "" or Environment("ErrorMsg") <> "" or blnRunFlag <>0 Then
																					Call fnInsertResult(strTestCaseID,strColumnValue,Chr(34) & strColumnValue & Chr(34) & " Component should be executed successfully",Chr(34) & strColumnValue & Chr(34) & " Component did not execute successfully","FAIL")
																					Execute"Set scrInstance = New SCR_Logoff"
																					scrInstance.Logoff
																					'Call CloseOpenBrowser()
																					print "Error:::" & strTestCAseID & " Exiting Due to Error in Component---" & strColumnValue
																					Exit For 'Since Comp has beeb failed
																		End If
																		Set rsSheet.ActiveConnection = Nothing
																		rsSheet.Close
															End If
												Else
															Set scrInstance = New SCR_Logoff
															scrInstance.Logoff
															'Call CloseOpenBrowser()
															Print "Error:::" & strTestCaseID & "Exiting Due to Error in Component ---" & strColumnValue
															'Since earlier Comp has been failed
												End If
											End If
									Next 
									'Call TC_ManualUpdate(Environment("FileName"))
									Select Case g_Flag
									Case 0
												strStatus = "Passed"
									Case 1
												strStatus = "Failed"
									End Select
									rsScenarios("ExecutionStatus").value = strStatus
									rsScenarios.Update
									Call fnSummaryInsertTestCase()
						'reset err variables
						Err.Clear
						blnRunFlag = 0
						Environment("ErrorMsg") = ""
			rsScenarios.MoveNext
			Wend
			Call fnSummaryCloseHtml(Environment.Value("Release"))
			If Environment.Value("MailTo")<>""Then
					Call fnSendSummarySnapshotEmail
			End If
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function fnLoadLibraries()
'			Call CloseOpenBrowser()
			strFrameworkFolderpath=Environment.Value("FrameworkFolderPath")
			Set objFso=CreateObject("Scripting.FileSystemObject")
			Set objFolder=objFSO.getFolder(Environment.Value("DataFolderPath"))
			Set sFiles=objFolder.Files
'			LoadfunctionLibrary strFrameworkFolderpath & "\Function Library\Verification.vbs"
'			LoadfunctionLibrary strFrameworkFolderpath & "\Function Library\HTML.vbs"
			For Each s in sFiles
						fname=s.name
						splitval=split(fname,".")
						If len(splitval(0))=5 then
								Environment("Module")=Mid(fname,4,2)
								Environment("DataFilePath")=objfolder.path & "\" & fname
'								LoadFunctionLibrary strFrameworkFolderpath & "\Object Dictionary\OBJ_" & environment.Value("Module")&".vbs"
'								LoadFunctionLibrary strFrameworkFolderpath & "\Screen Library\SCR_" & environment.Value("Module")&".vbs"
								Call CallMainScript
						End If
			Next
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function GetRecordSetValue(rstRecordset,rstRecordField)
			'On Error Resume Next
			GetRecordSetValue=Trim(rstRecordset(rstRecordField).value)
			'GetRecordSetValue=Trim(SubRecordSetValue(rstRecordset,rstRecordField))
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function SubRecordSetValue(rstRecordset,rstRecordField)
			'On Error Resume Next
			SubRecordSetValue=replace(rstRecordset(rstRecordField).value & "",chr(0),chr(26))
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function GetLocalRecordSet()
			'OnErrorResumeNext
			Dim adoRecordSet
			Set adoRecordSet=CreateObject("ADODB.Recordset")
			Set adoRecordSet.ActiveConnection=nothing
			adoRecordSet.CursorLocation=adUseClient
			adoRecordSet.Locktype=adLockBatchOptimistic
			Set GetLocalRecordSet=adoRecordSet
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function OpenDBConnection(strDBFormat,strHost,strDSN)
			'OnErrorResumeNext
			Dim adoConn,strConn
			strFilename=Environment.value("DataFilePath")
			Set adoConn=Createobject("ADODB.Connection")
			Select Case Ucase(trim(strDBFormat))
			Case "XLS"
						strConn="DRIVER={Microsoft Excel Driver (*.xls)};DBQ="&strFileName &";Readonly=False"
			End Select
			adoConn.Open strConn
			Set OpenDBConnection=adoConn
End Function


'==================================================================================================
'
'
'==================================================================================================
Function GetContentsFromDB(strFileName,strSQLStatement)
			'OnErrorResumeNext
			Dim objCon,objAdRs
			Set objAdCon=OpenDBConnection("XLS","","")
			Set objAdRs=GetLocalRecordSet()
			If Err<>0 then
							Reporter.ReportEvent micfail,"Create connection","[Connection]Error has Occured.error:"&Err
							Set obj_UDF_getRecordSet=nothing
                			Exit Function
			End If
            objAdRs.Open strSQLStatement,objAdCon,1,3
			If Err<>0 then
							Reporter.ReportEvent micfail,"Create connection","[Connection]Error has Occured.error:"&Err
							Set obj_UDF_getRecordSet=nothing
							Exit Function
			End If
			Set GetContentsFromDB=objAdRs
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function GetDisconnectedRecordSet(adoConn)
			'OnErrorResumeNext
			Dim adoRecordSet
			Set  adoRecordSet = GetLocatRecordSet()
			adoRecordSet.ActiveConnection = Nothing
			adoConn.Close
			Set adoConn = Nothing
			Set GetDisconnectedRecordSet = adoRecordSet
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function GetOpenBrowsers()
			'OnErrorResumeNext
			Dim descBrowser
			Set  descBrowser=Description.Create()
			descBrowser("micclass").value="browser"
			Set GetOpenBrowsers=Desktop.ChildObjects(descBrowser)
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function BrowserInstance(intIndex)
			'OnErrorResumeNext
			If intIndex<1 Then
					intIndex=1
			End If
			Set  BrowserInstance= Browser("Creationtime:=" &intIndex-1)
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function BrowserPage()
	'OnErrorResumeNext
	Set BrowserPage=BrowserInstance(GetOpenBrowsers().Count).Page("micclass:=Page")
	'Set BrowserPage=Browser("hwnd:=strHwnd").Page("micclass:=Page")
End Function

'==================================================================================================
'
'
'==================================================================================================
Function TC_ManualUpdate(strAttachmentFilePath)
			'OnErrorResumeNext
			'-Declare Variables-
			Dim TestFact,TestSetF,TSTestF,RunF,testsetitem,testitem,testInstance,att,rtt,testRunitem,objItem,Ist1,newinstance,item1,Ist3,TestFilter,strFile,TestList,oConn
			Dim intcounter,arrsAttachment,arrsAttachmentDesc,strStatus
'			Set oConn=CreateObject("TDApiOLe80.TDConnection")
'			oConn.InitConnectioEx Environment.Value("QC_URL")
'			oConn.Login  Environment.Value("QCID"),strQCPassword
'			'oConn.Login  Environment.Value("QCID"), Environment.Value("QCPassword")'strQCPassword
'			oConn.Connect   Environment.Value("QC_Domain"),  Environment.Value("QC_project")
			Set oConn = QCUtil.QCConnection
			intTestID=  Environment.Value("TestID")
			intTestSetID=  Environment.Value("TestSetID")
			'-Initiate variables-
			newinstance = 1
			'-Status Update
			Select Case g_Flag
			Case 0
						strStatus="passed"
			Case 1
						strStatus="Failed"
			Case Else
						strStatus="Failed"
			End Select
			rsScenarios("ExecutionStatus").value=strStatus
			rsScenarios.Update
			'-To Check If The TestCase Id is Provided in the Datasheet-
			If isnumeric(intTestID)=False Then
							Reporter.ReportEvent micfail,"Update Test Case:" &intTestId & "Failed","The TestCase Id is Null" 
							Call fnInsertResult(Environment("TestCaseID"),"Update Test Case:"&Chr(34)&intTestID &chr(34),"Test Case ID Should Be provided in the datasheet","The Test  ID is null","FAIL")
                			Exit Function
			End if
            '-To Check If The TestCase Id is Provided in the Datasheet-
			if isnumeric(intTestID)=False Then
							Reporter.ReportEvent micfail,"Update Test Case:" &intTestId & "Failed","The TestCase Id is Null" 
							Call fnInsertResult(Environment("TestCaseID"),"Update Test Case:"&Chr(34)&intTestID &chr(34),"Test Case ID Should Be provided in the datasheet","The Test  ID is null","FAIL")
							Exit Function
			End if
			'-Check for the QC connection-
            if oConn.Connected Then
							'=Setup of the QCobjects-
							Set TestFact=oConn.TestFactory
							Set TestSetF=oConn.TestSetFactory
							Set TSTestF=oConn.TSTestSetFactory
							Set RunF=oConn.RunFactory
							Set Testitem=TestFact.item(intTestID)
							Set TestSetitem=TestSetF.item(intTestSetID)
							Set Testinstance=TestSetitem.TSTestFactory
							
							'-Check if the Test Case exists in QC Test Plan-
							
							Set TestFilter=TestFact.Filter
							TestFilter("TS_TEST_ID")=intTestId
							Set TestList=TestFact.NewList(TestFilter.Text)
							if TestList.Count=0 Then
										TC_ManualUpdate=False
										Exit Function
							End if
							'-Setting the child objects for the testcases that needs to be run in QC-
							Set Ist1=testInstance.NewList("")
							For Each objItem In Ist1
										if cstr(objitem.field("TC_TEST_ID")) =cstr(inttestID)Then
													newInstance=0
													Set ist3=testInstance.newlist(objitem.id) 
													For Each item1 in ist3
																Set att=item1
													Next
													Exit for
										End If
							Next
							'-if the test case does not exist in the test set,the test case is added to the test Set-
							if newInstance=1 Then
										Set att=testInstance.additem(null)
										att.Field("TC_TEST_ID")=testitem
										att.post
							End If
							Set testRunItem=att.RunFactory
							Set rtt=testRunitem.additem(testRunitem.UniqueRunName)
							'-set the status of the overall run(for the test case)-
							rtt.Status=strStatus
							rtt.Post
							rtt.CopyDesignSteps
							rtt.Post
							Set strFille=rtt.Attachments.additem(null)
							strFile.FileName=strAttachmentFilePath & arrAttachment(intcounter)
							strFile.type=1
							strFile.Description="ExecutionReport of Testcase"
							strFile.Post
							strFile.Refresh
							rtt.Post
							Reporter.ReportEvent micfail,"TC_ManualUpdate " &intTestId & "Failed","The TestCase  status is updated in the test lab"
							Call fnInsertResult(Environment("TestCaseID"),"TC_ManualUpdate:"&Chr(34)&intTestID &chr(34),"Test Case  Status should be uploaded to QC","Test Case Status uploaded to QC","PASS" )
			Else
							Reporter.ReportEvent micfail,"TC_ManualUpdate " &intTestId & "Failed","The TestCase  status could not be completed due to QC Time out" 
							Call fnInsertResult(Environment("TestCaseID"),"TC_ManualUpdate:"&Chr(34)&intTestID &chr(34),"Test Case Test Case  Status should be uploaded to QC","Test Case Status should be uploaded to QC Time out","FAIL")
			End If
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function fncheckforErrors()
            'OnErrorResumeNext
			fncheckforErrors=True
			Do
				Wait 1
			Loop until objParent.Object.readyState="complete"'''Need to write function for 
			
			strErrorText = ""
			Set odes=description.create
			odes("micclass").value="WebElement"
			odes("class").value="message"
			
			Set Tblcnt = objParent.ChildObjects(odes)
			
			If Tblcnt.Count > 0 then
					For i = 0 To Tblcnt.Count-1
							str = Tblcnt(i).getroproperty("innertext")
							strErrorText = strErrorText & "_" & str
					Next
					Environment("ErrorMsg") = strErrorText
					fncheckforErrors = False
			End If
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function LaunchApp()
			REM 'On Error Resume Next
			REM 'Call fnDeleteCookies()
			REM 'Set oLogin=New UI_Login
			 strUrl=Environment.Value(""&Environment.Value("Region")&"")
			 Systemutil.Run "iexplore.exe",strUrl
			 'wait 2
End Function


''==================================================================================================
''
''
''==================================================================================================
'Public Function LaunchApp()
'			'On Error Resume Next
'			'Call fnDeleteCookies()
'			'Set oLogin=New UI_Login
'			Dim IE
'			Set IE = CreateObject("InternetExplorer.Application")
'			IE.Visible = True
'			IE.Navigate Environment.Value(""&Environment.Value("Region")&"")
'			strHwnd = IE.hwnd
'			Set IE = Nothing
'End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function CloseOpenBrowser()
			'On Error Resume Next
			Set objBro = Description.Create()
			objBro("micclass").Value = "Browser"
			Set objBrows = Desktop.ChildObjects(objBro)
			For i=0 to objBrows.count-1
							'Do not close the Browser Name "QualityCenter"
							If(objBrows(i).GetROProperty("title")<>"QualityCenter") and (IsEmpty(objBrows(i).GetROProperty("title")) = False)Then
										objBrows(i).close
							End If
			Next
'			If Browser("micclass:=browser").Dialog("text:=Windows Internet Explorer").WinButton("text:=OK").Exist(2) Then
'						Browser("micclass:=browser").Dialog("text:=Windows Internet Explorer").WinButton("text:=OK").Click
'			End If
'			If Browser("micclass:=browser").Page("micclass:=page").WebElement("html id:=button-1005-btnInnerEl").Exist(2) Then
'						Browser("micclass:=browser").Page("micclass:=page").WebElement("html id:=button-1005-btnInnerEl").Click
'			End if
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function SetPageValueDict(oDict, obj, strKeyName, strValue)
				'On Error Resume Next
				Environment("ErrorMsg") = ""
				blnDisabled = True 'Initialize it assuming the object is enabled
				'Err.Clear
				strClass = UCASE(Left(strKeyName, 3))
				intLength = Len(strKeyName)
				intChars = intLength - 3
				strDTColName = Right(strKeyName, intChars)
				Set ObjectName = obj.Object(oDict, strKeyName)
				'ObjectName.HighLight
				ActionData = Trim(strValue)
				If IsNull(ActionData) Then
							SetPageValueDict = True
							Exit Function
				End If
				blnRetVal = oVerify.fnVerifyobjectExist(ObjectName, strKeyName, strMsg)
				If not blnRetVal Then
								blnRunFlag = 1
								SetPageValueDict = False
								Exit Function
				End If
'				If strClass = "PGE" or strClass = "LNK" or strClass = "ELE" or strClass = "IMG" Then
'							'Skip
'				Else
'							blnDisabled = cBool(ObjectName.WaitProperty("disabled",False,Environment.Value("ObjectTimeOut")))
'				End If
				If Err.Number <> 0 Then
							Environment("ErrorMsg") = Err.Description
							Call fnInsertResult(Environment("TestCaseID"),Chr(34) & strKeyName & Chr(34), "Object Should be enabled to perform action", Err.Description, "FAIL")
							blnRunFlag = 1
							SetPageValueDict = False
							Exit Function
				End If
				If blnDisabled Then
								Select Case strClass
								Case "EDT"
											ObjectName.Object.value =  Trim(ActionData)
											ObjectName.Click
											Call fnInsertResult(Environment("TestCaseID"), strColumnValue, "System Should Perform Action on " & Chr(34) & strDTColName & Chr(34) & "as<b>" & Chr(34)& ActionData & Chr(34) & "</b> in" ,"System Performed the Action on " & Chr(34) & strDTColName & Chr(34),"DONE")
								Case "BTN","ELE","IMG","LNK"
											ObjectName.Click
											blnRetVal = fnCheckForErrors()
											Call fnInsertResult(Environment("TestCaseID"), strColumnValue, "System Should Perform Action on " & Chr(34) & strDTColName & Chr(34),"System Performed the Action on " & Chr(34) & strDTColName & Chr(34),"DONE")
								Case "CHK","RAD"
											ObjectName.Click
											blnRetVal = fnCheckForErrors()
											Call fnInsertResult(Environment("TestCaseID"), strColumnValue, "System Should Perform Action on " & Chr(34) & strDTColName & Chr(34),"System Performed the Action on " & Chr(34) & strDTColName & Chr(34),"DONE")
								Case "BLI"			
											ObjectName.Object.Click
											wait(1)
											Set Ob = Browser("micclass:=Browser").Page("micclass:=Page")
											If ob.Link("html tag:=A","innertext:=" & ActionData).Exist(Environment.Value("SyncTimeOut")) Then
												ob.Link("html tag:=A","innertext:=" & ActionData).Click
												blnRetVal = fnCheckForErrors()
												Call fnInsertResult(Environment("TestCaseID"), strColumnValue, "System Should Perform Action on " & Chr(34) & strDTColName & Chr(34) & "as<b>" & Chr(34)& ActionData & Chr(34) & "</b> in" ,"System Performed the Action on " & Chr(34) & strDTColName & Chr(34),"DONE")
											Else
												blnRetVal = False
											End If
								Case "PGE"
								'Skip
								Case "LST"			
											ObjectName.Object.Click
											wait(1)
											If ObjParent.WebElement("html tag:=LI","innertext:=" & ActionData).Exist(Environment.Value("SyncTimeOut")) Then
												ObjParent.WebElement("html tag:=LI","innertext:=" & ActionData).Click
												blnRetVal = fnCheckForErrors()
												Call fnInsertResult(Environment("TestCaseID"), strColumnValue, "System Should Perform Action on " & Chr(34) & strDTColName & Chr(34) & "as<b>" & Chr(34)& ActionData & Chr(34) & "</b> in" ,"System Performed the Action on " & Chr(34) & strDTColName & Chr(34),"DONE")												
											Else
												blnRetVal = False
											End If
								Case Else
												ObjectName.Click
												blnRetVal = fnCheckForErrors()
												Call fnInsertResult(Environment("TestCaseID"), strColumnValue, "System Should Perform Action on " & Chr(34) & strDTColName & Chr(34),"System Performed the Action on " & Chr(34) & strDTColName & Chr(34),"DONE")
								End Select
				Else
								blnRunFlag = 1
								SetPageValueDict = False
								Call fnInsertResult(Environment("TestCaseID"), Chr(34) & strDTColName & Chr(34), "Verification " & Chr(34) & strDTColName & Chr(34) & " should be disabled", Chr(34) & strDTColName & Chr(34) & " is disabled","FAIL")
				End If
				'Error Handling
				If Err.Number <> 0 or Environment("ErrorMsg") <>"" or blnRetVal = False Then
								blnRunFlag = 1
								SetPageValueDict = False
								Call fnInsertResult(Environment("TestCaseID"), Chr(34) & strDTColName & Chr(34), "System Should Perform Action on " & Chr(34) & strDTColName & Chr(34) & "as<b>" & Chr(34)& ActionData & Chr(34)& "</b> in" ,"System Failed to Perform the Action on" & Chr(34) & strDTColName & Chr(34) & "-->" & Err.Description &"-->"& Environment("ErrorMsg"),"FAIL")
				Else
								SetPageValueDict = True
				End If

End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function ReadPageValueDict(oDict, obj, strKeyName,strRow,strCol)
			'On Error Resume Next
			Environment("ErrorMsg") = ""
			Err.Clear
			strClass = UCASE(Left(strKeyName, 3))
			intLength = Len(strKeyName)
			intChars = intLength - 3
			strDTColName = Right(strKeyName, intChars)
			Set Objectname = obj.Object(oDict, strKeyName)

			blnRetVal = oVerify.fnVerifyobjectExist(ObjectName, strKeyName, strMsg)			
			'blnRetVal = oVerify.fnVerifyobjectExist(oDict, obj, strKeyName, strMsg)
			If not blnRetVal then
						blnRunFlag = 1
						ReadPageValueDict = False
						Exit Function
			End If
			
			If strClass = "EDT" or strClass = "LST" or strClass = "RGD" Then
						strValue = ObjectName.GetROProperty("Value")
			End If
			If strClass = "ELE" Then
						strValue = ObjectName.GetROProperty("innertext")
			End If
			If strClass = "CHK" Then
						strValue = ObjectName.GetROProperty("checked")
			End If
		
			If strClass = "TBL" Then
						strValue = Trim(ObjectName.GetCellData(strRow,strCol))
			End If
			'Error Handling
			If Err.Number <> 0  Then
						Environment("ErrorMsg") = Err.Description
						ReadPageValueDict = False
						Call fnInsertResult(Environment("TestCaseID"),"Read " & Chr(34) & strDTColName & Chr(34), "System should read " & Chr(34) & strDTColName & Chr(34) & "from Application", "Unable to Read the value: " & Chr(34)& strDTColName & Chr(34) & "from the Application"& Err.Description ,"FAIL")
			Else
						ReadPageValueDict = strValue
						'Call fnInsertResult(Environment("TestCaseID"),"Read value from " & Chr(34) & strDTColName & Chr(34), "System should read value from " & Chr(34) & strDTColName & Chr(34) & "from Application", "Read value is : " & Chr(34) & strValue & Chr(34) & " from " & Chr(34) & strDTColName & Chr(34) & " from the Application","PASS")
			End If
End Function 

'==================================================================================================
'
'
'==================================================================================================
Public Function fnOpenBrowser()
			'On Error Resume Next
			Set oDesc = Description.Create
			oDesc("micclass").Value = "Frame"
			oDesc("html tag").Value = "IFRAME"
			Set oFrames = BrowserPage().ChildObjects(oDesc)
			For i=0 to oFrames.Count-1
						If oFrames(i).GetROProperty("title")<>"" or err.description Then
									Set oDesc1 = Description.Create
									oDesc1("micclass").Value = "WebButton"
									oDesc("name").Value = "Exit"
									Set oButtons = oFrames(i).ChildObjects(oDesc1)
									For j=0 to oButtons.Count-1
												oButtons(j).Click
									Next
						End If
			Next
End Function

'==================================================================================================
'
'
'==================================================================================================
Function fnGetDTColName(strObjName)
			intLength = Len(strObjName)
			intChars = intLenght - 3
			strGetDTColName = Right(strObjName, intChars)
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function fnGenerateUniqueShortName()
			strDateTime = Now
			strDateTime = Repalce(strDateTime,Space(1),Space(0))
			intLenght = Len(strDateTime)
			strShortName = Left(strDateTime,intLength-2)
			fnGenerateUniqueShortName = strShortName
End Function
