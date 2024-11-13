'###############################################################################################################
'Test Script Name                          :  SSO_Login
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Login
'Test Case Name                            :  SSO_Login
'Author                                    :  Anil Kumar V
'Designed Date                             :  11 Feb 2015
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="SSO_Login"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="SSO_Login"
                StrStartTime=now
                StrBeginTimer=Timer

                strProjectName = environment("ProjectName")
                environment("strOSName") = Environment.Value("OS")
                
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" or environment("strOSName") = "Microsoft Windows 2012" Then
                                strInputCheck = "Users"
                ElseIf environment("strOSName") = "Microsoft Windows XP Workstation" Then
                                strInputCheck = "DOCUME~1"
                Else
                                Reporter.ReportEvent micFail, "Operating System Issue", "Host OS is not matching"
                                ExitAction "Fail: OS Issue - Host OS is not matching to XP / Vista / Windows 7"
                End If

'***************************************End ******************************************************************************

'***************************************Reading Input path and assigning it to respective variable***************
                environment("flag")=0
                strInputPath = Environment.Value("TestDir")
				strValueCheck = instr(1, strInputPath, strInputCheck)
                If strValueCheck <> 0 Then
					Set qtpApp = CreateObject("QuickTest.Application")
					qtpApp.Folders.RemoveAll
					qtpApp.Folders.Add ("[QualityCenter] Subject\Applications\248240:SSO\" & environment("ProjectName") & "\DriverData"),1
					strDriverDataFileName = environment("ProjectAbb") & "DriverData.xml"
					strConfigFilePath = PathFinder.Locate(strDriverDataFileName)
					environment("QC_DetailedRepPath")=strConfigFilePath
					Environment.LoadFromFile(strConfigFilePath)
					strDetailedReportFolder = "UFT"
					strDetailedReportPath = environment("QC_DetailedRepPath")
					strDetailedReportPath = left(strDetailedReportPath, instr(1, strDetailedReportPath, strDetailedReportFolder)-1)
					strDetailedReportPath = strDetailedReportPath & strDetailedReportFolder & "\"
					environment("strDetailedReportPath") = strDetailedReportPath
					strSumResultStatus="PASS"
					environment("flag")=1
					ExecuteFile environment("QC_GlobalVariablePath") &"\"& environment("ProjectAbb") & "GlobalVariables.vbs"
					Call QC_Associate_Utilities
                Else
				
					strInputPath = left(strInputPath, instr(1, strInputPath, strProjectName)-1)
					environment("strInputPath")=strInputPath
					strSumResultStatus="PASS"
					strGlobalVariablePath=environment("strInputPath") & environment("ProjectName") & "\CommonLibrary\GlobalVariables\" & environment("ProjectAbb") & "GlobalVariables.vbs"					
					ExecuteFile strGlobalVariablePath					
					Environment.LoadFromFile environment("strEnvPath")					
                    Call Associate_Utilities					
                End If
  
''***************************************End of reading input path******************************************************

				'Step -3 : Initialize application Base state
                Call Initialize_AppBaseState
				User= Environment("UserName")
               Environment("Browser_chrome") = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
				'Invoke Browser and Launch application
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" Or environment("strOSName")="Microsoft Windows XP Workstation" Then
                   
                   If Environment("Browser_Name")="Internet_Explorer" Then
                   		Call Func_InvokeBrowser(Environment("Browser_IE"),Environment("URL"))
                   	ElseIf Environment("Browser_Name")="Firefox" Then
                   		Call Func_InvokeBrowser(Environment("Browser_FF"),Environment("URL"))
                   	else
                   		 Call Func_InvokeBrowser(Environment("Browser_chrome"),Environment("URL"))
                   End If                                
                Else
                    Call Func_InvokeBrowser(Environment("Browser"),Environment("URL"))
                End If
				'On Error Resume Next
				
                If environment("flag")=1 Then
					strDataFile = strDataFile_QC
					environment("strDataFile")=strDataFile
				End If

'#################################### TEST  CASE VALIDATION STARTS HERE ##################################################################################
                
                DataTable.ImportSheet environment("strDataFile"), "Global", "Global"
                environment("row_count") = DataTable.GetSheet("Global").GetRowCount
				
				If environment("row_count") = 0 Then
						Func_CreateRow strDtlRptFileName,environment("strTestCaseName"),"Row count - Data table row count is zero or not exists, please input valid data in the data table","Fail",now
						strSumResultStatus="FAIL"
						Call Func_WriteSummaryReportAndUploadToQC()
						Reporter.ReportEvent micFail, "Input Data", "Row count - Data table row count is zero or not exists, please input valid data in the data table"
						ExitAction "Fail: Row count - Data table row count is zero or not exists, please input valid data in the data table"
				Else

					For i = 1 to environment("row_count")
                            																	

							Do Until Obj_Wbtn_Sgn.exist(1)
							Loop
'							If Browser("Problem loading page").Page("Problem loading page").WebElement("Server not found").Exist(1) then
'							 	Func_CreateRow strDtlRptFileName,strTestCaseName, "Login failed due to: ","Fail",now
'								Call Func_WriteSummaryReportAndUploadToQC()
'								ExitAction "Fail: Login failed "
'							End IF

							Call Func_Object_Exists(Obj_Wbtn_Sgn)
							Call Func_Do_Action(Obj_Wbtn_Sgn,"TRUE","")
							
							obj_pg_Home.Refreshobject
							Call Func_Object_Exists(Obj_Wbedt_Un)
							Call Func_Do_Action(Obj_Wbedt_Un,Environment("User"),WebEdit)
							
							Call Func_Object_Exists(Obj_Wbedt_Pwd)			
							Call Func_Do_Action(Obj_Wbedt_Pwd,Environment("Pwd"),WebEdit)
							
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")

'							Set sgn_btn=Browser("Welcome - HPE Software").Page("Welcome - HPE Software").WebButton("sign")
'						
'							Call Func_Object_Exists(sgn_btn)
'							Call Func_Do_Action(sgn_btn,"TRUE","")
'							
'							Call Func_Object_Exists(Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").WebEdit("username"))
'							Call Func_Do_Action(Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").WebEdit("username"),Environment("User"),WebEdit)
'							
'							Call Func_Object_Exists(Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").WebEdit("password"))			
'							Call Func_Do_Action(Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").WebEdit("password"),Environment("Pwd"),WebEdit)
'							
'							Call Func_Object_Exists(Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").Link("Sign in"))
'							Call Func_Do_Action(Browser("Welcome - HPE Software").Page("Sign in | HPE® Official").Link("Sign in"),"TRUE","")

							If obj_Wbele_err.exist(3) Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Login failed due to: " & obj_Wbele_err.getroproperty("innertext"),"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail: Login failed "
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","Pass",now							
							End If									
							
							Call Func_Object_Exists(Obj_Lnk_Sgt)			
							Call Func_Do_Action(Obj_Lnk_Sgt,"TRUE","")
							
							Func_CreateRow strDtlRptFileName,strTestCaseName,"sucessfully Logged out from OSP Application","Pass",now
														
						next

				end if

'#################################### End of Test Case Validation #########################################################################################

'''***********************Writting final Summary Result for the test case in the summary report file & Uploading the Detailed test execution report to Quality Center Detailed result folder*********
Call Func_WriteSummaryReportAndUploadToQC()

'************************End of Summary Report and Uploading detailed report into Quality Center******************************************************************************************************

''****************************************************Script End***************************************************************************************************************************************

''=============	Browser and Page Object	==========================
'Set obj_pg_Home=Browser("CreationTime:=0").Page("Micclass:=Page")
'
'
''=============	Child Objects	==========================
'Set Wbtn_Sgn=Description.Create
'Wbtn_Sgn("name").Value="Sign In "
'Wbtn_Sgn("html tag").Value="A"
'Wbtn_Sgn("role").Value="menuitem"
'Set Obj_Wbtn_Sgn=obj_pg_Home.Link(Wbtn_Sgn)
'Obj_Wbtn_Sgn.Highlight


'Set Wbtbl=Description.Create
'Wbtbl("html tag").Value="DIV"
'Wbtbl("role").Value="grid"
'Wbtbl("index").Value=1
'Set obj_Wbtbl=Br_Pg.WebTable(Wbtbl)
'obj_Wbtbl.Highlight

'Set Wbtbl=Description.Create
'Wbtbl("class").Value="ui-grid-row ng-scope"
'Wbtbl("role").Value="row"
'Wbtbl("index").Value=1
'Set obj_Wbtbl=Br_Pg.WebElement(Wbtbl)
'obj_Wbtbl.Highlight


'
'MsgBox obj_Wbtbl.ChildItem(2,7,"WebElemenet",0).Exist
'MsgBox val.getElementsByTagName("SPAN").Item(0).textContent
'rc= obj_Wbtbl.RowCount
'
'For i = 1 To rc Step 1
'	CC=obj_Wbtbl.ColumnCount(i)
'	For j = 1 To CC Step 1
'		If Err.Number <> 0 Then
'			Print Err.Description
'			Err.Clear
'		Else
'			Print obj_Wbtbl.GetCellData(i,j)
'		End If
'	Next
'Next
'On Error Goto 0

''Author: Tarun Lalwani
'Website: http://KnowledgeInbox.com
''Description: GetWebTableFromElement functions uses DOM to find the parent TABLE which 
''				contains the current object and returns the corresponding WebTable 
''Params:
'' @pObject: Any object inside the table
'Public Function GetWebTableFromElement(ByVal pObject)
'    Dim oTable, oParent
'
'	'Get the parent DOM node
'	Set oTable = pObject.Object.parentNode
'
'	While oTable.tagName <> "DIV"
'		Set oTable = oTable.parentNode
'	Wend
'
'	'Get the parent Test Object
'    Set oParent = pObject.GetTOProperty("parent")
'
'	'If oTable is nothing it means we didn't get any parent with tag TABLE
'    If oTable Is Nothing Then
'        Set GetWebTableFromElement = Nothing
'    Else
'		'Get the element
'        Set GetWebTableFromElement = oParent.WebTable("source_Index:=" & oTable.sourceIndex)
'    End If
'End Function
'
'
''Author: Tarun Lalwani
''Website: http://KnowledgeInbox.com
''Description: GetWebTableRowFromElement functions uses DOM to find the parent TABLE which 
''				contains the current object and returns the corresponding WebElement for the Row 
''Params:
'' @pObject: Any object inside the table
'Public Function GetWebTableRowFromElement(ByVal pObject)
'    Dim oTable, oParent
'
'	'Get the parent DOM node
'	Set oTable = pObject.Object.parentNode
'
'	While oTable.tagName <> "TR"
'		Set oTable = oTable.parentNode
'	Wend
'
'	'Get the parent Test Object
'    Set oParent = pObject.GetTOProperty("parent")
'
'	'If oTable is nothing it means we didn't get any parent with tag TABLE
'    If oTable Is Nothing Then
'        Set GetWebTableRowFromElement = Nothing
'    Else
'		'Get the element
'        Set GetWebTableRowFromElement = oParent.WebElement("source_Index:=" & oTable.sourceIndex)
'    End If
'End Function
'
''Author: Tarun Lalwani
''Website: http://KnowledgeInbox.com
''Description: GetWebTableCellFromElement functions uses DOM to find the parent TABLE which 
''				contains the current object and returns the corresponding WebElement for the Cell 
''Params:
'' @pObject: Any object inside the table
'Public Function GetWebTableCellFromElement(ByVal pObject)
'    Dim oTable, oParent
'
'	'Get the parent DOM node
'	Set oTable = pObject.Object.parentNode
'
'	While oTable.tagName <> "TD"
'		Set oTable = oTable.parentNode
'	Wend
'
'	'Get the parent Test Object
'    Set oParent = pObject.GetTOProperty("parent")
'
'	'If oTable is nothing it means we didn't get any parent with tag TABLE
'    If oTable Is Nothing Then
'        Set GetWebTableCellFromElement = Nothing
'    Else
'		'Get the element
'        Set GetWebTableCellFromElement = oParent.WebElement("source_Index:=" & oTable.sourceIndex)
'    End If
'End Function
