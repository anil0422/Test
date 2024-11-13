'###############################################################################################################
'Test Script Name                          :  SSO_Navigate_All_Pages
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Navigate_All_Pages
'Test Case Name                            :  SSO_Navigate_All_Pages
'Author                                    :  Anil Kumar V
'Designed Date                             :  26 Apr 2015
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="SSO_Navigate_All_Pages"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="SSO_Navigate_All_Pages"
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
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" Then
'                    Call Func_InvokeBrowser(Environment("Browser_IE"),Environment("URL"))
                    Call Func_InvokeBrowser(Environment("Browser_FF"),Environment("URL"))
'                    Call Func_InvokeBrowser(Environment("Browser_chrome"),Environment("URL"))
                Else
                    Call Func_InvokeBrowser(Environment("Browser"),Environment("URL"))
                End If

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
                            																	
'							Environment("User") = DataTable.GetSheet("Global").GetParameter("User").ValueByRow(i)
'							Environment("Pwd") = DataTable.GetSheet("Global").GetParameter("Pwd").ValueByRow(i)
'							Environment("SAID")= DataTable.GetSheet("Global").GetParameter("SAID").ValueByRow(i)'"103404038965"'103724929751,000000000003,000000000004,102552308463,103404038965'DataTable.GetSheet("Global").GetParameter("SAID").ValueByRow(i)
							
							
'							obj_pg_Home.Sync
							Do Until Obj_Wbtn_Sgn.exist(1)
							Loop
							
							Call Func_Object_Exists(Obj_Wbtn_Sgn)
							Call Func_Do_Action(Obj_Wbtn_Sgn,"TRUE","")
							
'							obj_pg_Home.Sync
			
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
								Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","DONE",now							
							End If
														
							Call Func_Object_Exists(Obj_lnk_chk_entl)			
							Call Func_Do_Action(Obj_lnk_chk_entl,"TRUE","")
							
							Do Until Obj_wbele_chk_entl.exist(1)
							Loop
							
							If Obj_wbele_chk_entl.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_wbele_chk_entl.GetRoproperty("innertext")&" page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Check Entitlement Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
													
							Call Func_Object_Exists(Obj_Lnk_Cnfg_itms)			
							Call Func_Do_Action(Obj_Lnk_Cnfg_itms,"TRUE","")
'						
							Do Until Obj_Wbele_Cnfg_Itms.exist(1)
							Loop
							
							If Obj_Wbele_Cnfg_Itms.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_Wbele_Cnfg_Itms.GetRoproperty("innertext")&" page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Configuration Items Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
															
							Call Func_Object_Exists(Obj_Lnk_Email_Notif)			
							Call Func_Do_Action(Obj_Lnk_Email_Notif,"TRUE","")
'							
							Do Until Obj_Wbele_email_notf.exist(1)
							Loop
							
							If Obj_Wbele_email_notf.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_Wbele_email_notf.GetRoproperty("innertext")&" page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Email notification Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If	
							
							Call Func_Object_Exists(Obj_Lnk_srvy_pre)			
							Call Func_Do_Action(Obj_Lnk_srvy_pre,"TRUE","")
'							
							Do Until Obj_wbele_srvy_pre.exist(1)
							Loop
							
							If Obj_wbele_srvy_pre.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_wbele_srvy_pre.GetRoproperty("innertext")&" page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Survey Preferences Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If	
							
							Call Func_Object_Exists(Obj_Lnk_flx_care)			
							Call Func_Do_Action(Obj_Lnk_flx_care,"TRUE","")
							
							Do Until Obj_wbele_crd_avl.exist(1)
							Loop
							
							If Obj_wbele_crd_avl.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Flexcare Credit Management page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Flexcare Credit Management Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
					
							Call Func_Object_Exists(Obj_Lnk_dashboard)			
							Call Func_Do_Action(Obj_Lnk_dashboard,"TRUE","")
							
							Do Until Obj_wbele_dashboard.exist(1)
							Loop
							
							If Obj_wbele_dashboard.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Dashboard page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display dashboard Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_Lnk_My_Prd)			
							Call Func_Do_Action(Obj_Lnk_My_Prd,"TRUE","")
							
							Do Until Obj_Wbele_prd.exist(1)
							Loop
							
							If Obj_Wbele_prd.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_Wbele_prd.GetRoproperty("innertext")&" page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display My products Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_lnk_Manuals)			
							Call Func_Do_Action(Obj_lnk_Manuals,"TRUE","")
							
							If Obj_wbele_srch_kng.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Manuals page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Manuals Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_lnk_patch)			
							Call Func_Do_Action(Obj_lnk_patch,"TRUE","")
					
							If Obj_wbele_srch_kng.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Patches page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Manuals Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_lnk_srch)			
							Call Func_Do_Action(Obj_lnk_srch,"TRUE","")
					
							If Obj_wbele_srch_kng.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Search Knowledge page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Manuals Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_lnk_chng_req)			
							Call Func_Do_Action(Obj_lnk_chng_req,"TRUE","")
					
							If Obj_wbele_srch_kng.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Change Requests page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Manuals Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_lnk_SR_dash)			
							Call Func_Do_Action(Obj_lnk_SR_dash,"TRUE","")
					
							If obj_Wbtn_Sbt_New.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Service Request page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Manuals Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_lnk_Prd_News)			
							Call Func_Do_Action(Obj_lnk_Prd_News,"TRUE","")
					
							If Obj_wbele_prd_nws.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_wbele_prd_nws.GetRoproperty("innertext")&" page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display product News Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_lnk_Sup_News)			
							Call Func_Do_Action(Obj_lnk_Sup_News,"TRUE","")
					
							If Obj_wbele_sup_nws.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_wbele_sup_nws.GetRoproperty("innertext")&" page displayed sucessfully" ,"DONE",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Suport News Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
							
							Call Func_Object_Exists(Obj_Lnk_Sgt)			
							Call Func_Do_Action(Obj_Lnk_Sgt,"TRUE","")
							
							Func_CreateRow strDtlRptFileName,strTestCaseName,"sucessfully Logged out from OSP Application","Pass",now
	
'   			End If							
						next

				end if

'#################################### End of Test Case Validation #########################################################################################

'''***********************Writting final Summary Result for the test case in the summary report file & Uploading the Detailed test execution report to Quality Center Detailed result folder*********
Call Func_WriteSummaryReportAndUploadToQC()

'************************End of Summary Report and Uploading detailed report into Quality Center******************************************************************************************************

''****************************************************Script End***************************************************************************************************************************************
