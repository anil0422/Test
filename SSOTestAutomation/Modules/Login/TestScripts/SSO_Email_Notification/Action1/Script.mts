'###############################################################################################################
'Test Script Name                          :  SSO_Email_Notification
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Email_Notification
'Test Case Name                            :  SSO_Email_Notification
'Author                                    :  Anil Kumar V
'Designed Date                             :  1 June 2015
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="SSO_Email_Notification"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="SSO_Email_Notification"
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
                
                DataTable.ImportSheet environment("strDataFile"), "OSP_Email_Notification", "Global"
                environment("row_count") = DataTable.GetSheet("Global").GetRowCount
				
				If environment("row_count") = 0 Then
						Func_CreateRow strDtlRptFileName,environment("strTestCaseName"),"Row count - Data table row count is zero or not exists, please input valid data in the data table","Fail",now
						strSumResultStatus="FAIL"
						Call Func_WriteSummaryReportAndUploadToQC()
						Reporter.ReportEvent micFail, "Input Data", "Row count - Data table row count is zero or not exists, please input valid data in the data table"
						ExitAction "Fail: Row count - Data table row count is zero or not exists, please input valid data in the data table"
				Else

					For i = 1 to environment("row_count")
                            																	
							Environment("Product") = DataTable.GetSheet("Global").GetParameter("Product").ValueByRow(i)
							Environment("Version") = DataTable.GetSheet("Global").GetParameter("Version").ValueByRow(i)
							Environment("OS_Sel")=DataTable.GetSheet("Global").GetParameter("OS_Sel").ValueByRow(i)
							Environment("Sub_Product")=DataTable.GetSheet("Global").GetParameter("Sub_Product").ValueByRow(i)
													
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
								Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","Pass",now							
							End If
							
							Call Func_Object_Exists(Obj_Lnk_Email_Notif)			
							Call Func_Do_Action(Obj_Lnk_Email_Notif,"TRUE","")
							
							If Obj_Wbele_email_notf.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_email_notf.GetRoproperty("innertext")&" page displayed sucessfully ","Pass",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName," Failed to display e-mail notification page","Fail",now	
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed: to display e-mail notification page"
							End If
							
							Call Func_Object_Exists(Obj_Lnk_Reg_doc)			
							Call Func_Do_Action(Obj_Lnk_Reg_doc,"TRUE","")
							
							Obj_Wbele_email_notf.Refreshobject
							If Obj_Wbele_email_notf.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_email_notf.GetRoproperty("innertext")&" page displayed sucessfully ","Pass",now
								
								Call Func_Object_Exists(Obj_wblst_reg_prd_sel)			
								Call Func_Do_Action(Obj_wblst_reg_prd_sel,Environment("Product"),WebList)
								
								wait(10)
								Call Func_Object_Exists(Obj_wblst_reg_ver_sel)			
								Call Func_Do_Action(Obj_wblst_reg_ver_sel,Environment("Version"),WebList)
								
'								wait(2)
'								Call Func_Object_Exists(Obj_wblst_reg_os_sel)			
'								Call Func_Do_Action(Obj_wblst_reg_os_sel,Environment("OS_Sel"),WebList)
'								
'								Call Func_Object_Exists(Obj_wblst_reg_Sub_sel)			
'								Call Func_Do_Action(Obj_wblst_reg_Sub_sel,Environment("Sub_Product"),WebList)
								
								Call Func_Object_Exists(obj_wbchk_doc_typ)	
								obj_wbchk_doc_typ.click								
'								Call Func_Do_Action(obj_wbchk_doc_typ,"TRUE","")
								
								Call Func_Object_Exists(obj_Wbtn_Regst)			
								Call Func_Do_Action(obj_Wbtn_Regst,"TRUE","")
							
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Register for document e-mail notification page not displayed sucessfully ","Fail",now
							End If
							
							If Obj_Wbele_thanks.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_thanks.GetRoproperty("innertext")&" page displayed sucessfully ","Pass",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Thanks page failed to display after register ","Fail",now
							End If
											
							Obj_Lnk_Email_Notif.Refreshobject											
							Call Func_Object_Exists(Obj_Lnk_Email_Notif)			
							Call Func_Do_Action(Obj_Lnk_Email_Notif,"TRUE","")
													
							Call Func_Object_Exists(Obj_Lnk_Reg_chg_req)			
							Call Func_Do_Action(Obj_Lnk_Reg_chg_req,"TRUE","")
							
							Obj_Wbele_email_notf.Refreshobject
							If Obj_Wbele_email_notf.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_email_notf.GetRoproperty("innertext")&" page displayed sucessfully ","Pass",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Register for change request e-mail notification page not displayed sucessfully ","Fail",now
							End If
							
							Obj_Lnk_Email_Notif.Refreshobject											
							Call Func_Object_Exists(Obj_Lnk_Email_Notif)			
							Call Func_Do_Action(Obj_Lnk_Email_Notif,"TRUE","")
													
							Call Func_Object_Exists(Obj_Lnk_Reg_ser_req)			
							Call Func_Do_Action(Obj_Lnk_Reg_ser_req,"TRUE","")
							
							Obj_Wbele_email_notf.Refreshobject
							If Obj_Wbele_email_notf.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_email_notf.GetRoproperty("innertext")&" page displayed sucessfully ","Pass",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Register for service request e-mail notification page not displayed sucessfully ","Fail",now
							End If
									
							Obj_Lnk_Email_Notif.Refreshobject											
							Call Func_Object_Exists(Obj_Lnk_Email_Notif)			
							Call Func_Do_Action(Obj_Lnk_Email_Notif,"TRUE","")
													
							Call Func_Object_Exists(Obj_Lnk_Reg_ser_req)			
							Call Func_Do_Action(Obj_Lnk_Reg_ser_req,"TRUE","")
							
							Obj_Wbele_email_notf.Refreshobject
							If Obj_Wbele_email_notf.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_email_notf.GetRoproperty("innertext")&" page displayed sucessfully ","Pass",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"delete e-mail registrations page not displayed sucessfully ","Fail",now
							End If
							
							Call Func_Object_Exists(Obj_Lnk_Sgt)			
							Call Func_Do_Action(Obj_Lnk_Sgt,"TRUE","")
							
							Func_CreateRow strDtlRptFileName,strTestCaseName,"sucessfully Logged out from SSO Application","Pass",now
	
'   			End If
							
						next

				end if

'#################################### End of Test Case Validation #########################################################################################

'''***********************Writting final Summary Result for the test case in the summary report file & Uploading the Detailed test execution report to Quality Center Detailed result folder*********
Call Func_WriteSummaryReportAndUploadToQC()

'************************End of Summary Report and Uploading detailed report into Quality Center******************************************************************************************************

''****************************************************Script End***************************************************************************************************************************************
'Set obj_pg_Home=Browser("CreationTime:=0").Page("Micclass:=Page")
'
'Set Lnk_Reg_ser_req=Description.Create
'Lnk_Reg_ser_req("name").Value="Register for service request e-mail notification"
'Lnk_Reg_ser_req("html tag").Value="A"
'Set Obj_Lnk_Reg_ser_req=obj_pg_Home.Link(Lnk_Reg_ser_req)
'Obj_Lnk_Reg_ser_req.Highlight
'Obj_Lnk_Reg_ser_req.Click
