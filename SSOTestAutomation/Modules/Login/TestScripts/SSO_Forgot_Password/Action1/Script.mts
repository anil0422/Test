'###############################################################################################################
'Test Script Name                          :  SSO_Forgot_Password
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Forgot_Password
'Test Case Name                            :  SSO_Forgot_Password
'Author                                    :  Anil Kumar V
'Designed Date                             :  12 Feb 2015
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="SSO_Forgot_Password"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="SSO_Forgot_Password"
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
					strDetailedReportFolder = "QTPro"
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
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" Or environment("strOSName")="Microsoft Windows XP Workstation"  Then
                   
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
                
                DataTable.ImportSheet environment("strDataFile"), "OSP_Forgot_Password", "Global"
                environment("row_count") = DataTable.GetSheet("Global").GetRowCount
				
				If environment("row_count") = 0 Then
						Func_CreateRow strDtlRptFileName,environment("strTestCaseName"),"Row count - Data table row count is zero or not exists, please input valid data in the data table","Fail",now
						strSumResultStatus="FAIL"
						Call Func_WriteSummaryReportAndUploadToQC()
						Reporter.ReportEvent micFail, "Input Data", "Row count - Data table row count is zero or not exists, please input valid data in the data table"
						ExitAction "Fail: Row count - Data table row count is zero or not exists, please input valid data in the data table"
				Else

					For i = 1 to environment("row_count")
                            																	
'							environment("Email_Address") = DataTable.GetSheet("Global").GetParameter("Email_Address").ValueByRow(i)
							environment("Security_Ans1") = DataTable.GetSheet("Global").GetParameter("Security_Ans1").ValueByRow(i)
							environment("Security_Ans2") = DataTable.GetSheet("Global").GetParameter("Security_Ans2").ValueByRow(i)
							environment("New_Pass") = DataTable.GetSheet("Global").GetParameter("New_Pass").ValueByRow(i)
							environment("Conf_Pass") = DataTable.GetSheet("Global").GetParameter("Conf_Pass").ValueByRow(i)

'							obj_pg_Home.Sync
							Do Until Obj_Wbtn_Sgn.exist(1)
							Loop
							
							Call Func_Object_Exists(Obj_Wbtn_Sgn)
							Call Func_Do_Action(Obj_Wbtn_Sgn,"TRUE","")
							
							Call Func_Object_Exists(Obj_Lnk_Frgt_Pwd)
							Call Func_Do_Action(Obj_Lnk_Frgt_Pwd,"TRUE","")
							
'							obj_pg_Home.Sync

							Call Func_Object_Exists(Obj_Wbedt_email)
							Call Func_Do_Action(Obj_Wbedt_email,Environment("User"),WebEdit)
																					
							Call Func_Object_Exists(Obj_lnk_Nxt)
							Call Func_Do_Action(Obj_lnk_Nxt,"TRUE","")
														
							Call Func_Object_Exists(Obj_Wbedt_Sec_Ans1)
							Call Func_Do_Action(Obj_Wbedt_Sec_Ans1,Environment("Security_Ans1"),WebEdit)
							
							Call Func_Object_Exists(Obj_Wbedt_Sec_Ans2)
							Call Func_Do_Action(Obj_Wbedt_Sec_Ans2,Environment("Security_Ans2"),WebEdit)
							
							Obj_lnk_Nxt.Refreshobject
							Call Func_Object_Exists(Obj_lnk_Nxt)
							Call Func_Do_Action(Obj_lnk_Nxt,"TRUE","")
							
							Call Func_Object_Exists(Obj_Wbedt_Pwd)
							Call Func_Do_Action(Obj_Wbedt_Pwd,"test@1234567",WebEdit)
							
							Call Func_Object_Exists(Obj_Wbet_CPwd)
							Call Func_Do_Action(Obj_Wbet_CPwd,"test@1234567",WebEdit)
							
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")
							
							If Obj_Lnk_mail_nf.exist(2) Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_Lnk_mail_nf.getroproperty("innertext")&" Change the testdata and try again. " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail: Login failed "
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Your new password has been established. as "&Environment("Conf_Pass"),"Pass",now							
							End If
							
							Call Func_Object_Exists(Obj_lnk_chg_pwd)
							Call Func_Do_Action(Obj_lnk_chg_pwd,"TRUE","")
							
							Call Func_Object_Exists(Obj_Wbedt_Cur_Pwd)
							Call Func_Do_Action(Obj_Wbedt_Cur_Pwd,"test@1234567",WebEdit)
													
							Call Func_Object_Exists(Obj_Wbet_Pwd)
							Call Func_Do_Action(Obj_Wbet_Pwd,Environment("Pwd"),WebEdit)
							
							Obj_Wbet_CPwd.RefreshObject
							Call Func_Object_Exists(Obj_Wbet_CPwd)
							Call Func_Do_Action(Obj_Wbet_CPwd,Environment("Pwd"),WebEdit)
							
							Call Func_Object_Exists(Obj_lnk_Sav_chg)			
							Call Func_Do_Action(Obj_lnk_Sav_chg,"TRUE","")
							
							Call Func_Object_Exists(Obj_lnk_Ctn_site)
							Call Func_Do_Action(Obj_lnk_Ctn_site,"TRUE","")							
							
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

