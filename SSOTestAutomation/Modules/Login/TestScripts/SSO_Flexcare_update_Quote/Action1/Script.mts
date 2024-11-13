'###############################################################################################################
'Test Script Name                          :  SSO_Flexcare_update_Quote
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Flexcare_update_Quote
'Test Case Name                            :  SSO_Flexcare_update_Quote
'Author                                    :  Anil Kumar V
'Designed Date                             :  26 Apr 2016
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="SSO_Flexcare_update_Quote"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="SSO_Flexcare_update_Quote"
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
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" Then
'                    Call Func_InvokeBrowser(Environment("Browser_IE"),Environment("URL"))
                    Call Func_InvokeBrowser(Environment("Browser_FF"),Environment("URL"))
'                    Call Func_InvokeBrowser(Environment("Browser_chrome"),Environment("URL"))
                Else
                    Call Func_InvokeBrowser(Environment("Browser"),Environment("URL"))
                End If
				'On Error Resume Next
				
                If environment("flag")=1 Then
					strDataFile = strDataFile_QC
					environment("strDataFile")=strDataFile
				End If

'#################################### TEST  CASE VALIDATION STARTS HERE ##################################################################################
                
                DataTable.ImportSheet environment("strDataFile"), "Flx_crdt", "Global"
                environment("row_count") = DataTable.GetSheet("Global").GetRowCount
				
				If environment("row_count") = 0 Then
						Func_CreateRow strDtlRptFileName,environment("strTestCaseName"),"Row count - Data table row count is zero or not exists, please input valid data in the data table","Fail",now
						strSumResultStatus="FAIL"
						Call Func_WriteSummaryReportAndUploadToQC()
						Reporter.ReportEvent micFail, "Input Data", "Row count - Data table row count is zero or not exists, please input valid data in the data table"
						ExitAction "Fail: Row count - Data table row count is zero or not exists, please input valid data in the data table"
				Else

					For i = 1 to environment("row_count")
                            Environment("id")=DataTable.GetSheet("Global").GetParameter("id").ValueByRow(i)	
							Environment("comments")=DataTable.GetSheet("Global").GetParameter("comments").ValueByRow(i)	                            

							Do Until Obj_Wbtn_Sgn.exist(1)
							Loop
'							If Browser("Problem loading page").Page("Problem loading page").WebElement("Server not found").Exist(1) then
'							 	Func_CreateRow strDtlRptFileName,strTestCaseName, "Login failed due to: ","Fail",now
'								Call Func_WriteSummaryReportAndUploadToQC()
'								ExitAction "Fail: Login failed "
'							End IF

							Call Func_Object_Exists(Obj_Wbtn_Sgn)
							Call Func_Do_Action(Obj_Wbtn_Sgn,"TRUE","")
							
							Call Func_Object_Exists(Obj_Wbedt_Un)
							Call Func_Do_Action(Obj_Wbedt_Un,Environment("User"),WebEdit)
							
							Call Func_Object_Exists(Obj_Wbedt_Pwd)			
							Call Func_Do_Action(Obj_Wbedt_Pwd,Environment("Pwd"),WebEdit)
						
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")
							
							If obj_Wbele_err.exist(3) Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Login failed due to: " & obj_Wbele_err.getroproperty("innertext"),"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail: Login failed "
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","Pass",now							
							End If									
							
							Call Func_Object_Exists(Obj_Lnk_flx_care)			
							Call Func_Do_Action(Obj_Lnk_flx_care,"TRUE","")
							
							If Obj_wbele_flx_err.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Flexcare page not displayed "&Obj_wbele_flx_err.GetRoProperty("innertext") ,"FAIL",now	
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail "
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Flexcare management page displayed sucessfully","Pass",now								
							End If
							
							
							
							Call Func_Object_Exists(Browser("Welcome - HPE Software").Page("Flexcare Credit Management").WebTable("+"))		
							quote=Browser("Welcome - HPE Software").Page("Flexcare Credit Management").WebTable("+").Getcelldata(1,1)							
							Func_CreateRow strDtlRptFileName,strTestCaseName, "Clicked on quote id "&quote &" to update"  ,"PASS",now
							Browser("Welcome - HPE Software").Page("Flexcare Credit Management").WebTable("+").ChildItem(1,1,"Link",1).click
							
						
							Call Func_Object_Exists(obj_wbtd_cmnts)
							Call Func_Do_Action(obj_wbtd_cmnts,Environment("comments"),WebEdit)
							
							Call Func_Object_Exists(Obj_Wbtn_sbt_qt)			
							Call Func_Do_Action(Obj_Wbtn_sbt_qt,"TRUE","")
							
							If Obj_wbele_qt_suc1.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Quote Updated Sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Failed to update quote" ,"FAIL",now	
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail "								
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

