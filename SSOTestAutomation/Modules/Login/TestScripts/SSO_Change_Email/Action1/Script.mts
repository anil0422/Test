'###############################################################################################################
'Test Script Name                          :  SSO_Change_Email
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Change_Email
'Test Case Name                            :  SSO_Change_Email
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
                environment("StrTestCaseID")="SSO_Change_Email"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="SSO_Change_Email"
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
                
                DataTable.ImportSheet environment("strDataFile"), "OSP_Change_Email", "Global"
                environment("row_count") = DataTable.GetSheet("Global").GetRowCount
				
				If environment("row_count") = 0 Then
						Func_CreateRow strDtlRptFileName,environment("strTestCaseName"),"Row count - Data table row count is zero or not exists, please input valid data in the data table","Fail",now
						strSumResultStatus="FAIL"
						Call Func_WriteSummaryReportAndUploadToQC()
						Reporter.ReportEvent micFail, "Input Data", "Row count - Data table row count is zero or not exists, please input valid data in the data table"
						ExitAction "Fail: Row count - Data table row count is zero or not exists, please input valid data in the data table"
				Else

					For i = 1 to environment("row_count")
                            																	
'							Environment("User_ID") = DataTable.GetSheet("Global").GetParameter("User_ID").ValueByRow(i)
'							Environment("Pwd") = DataTable.GetSheet("Global").GetParameter("Pwd").ValueByRow(i)
							Environment("New_User_ID") = DataTable.GetSheet("Global").GetParameter("New_User_ID").ValueByRow(i)
							
'							obj_pg_Home.Sync
							Do Until Obj_Wbtn_Sgn.exist(1)
							Loop
							
							Call Func_Object_Exists(Obj_Wbtn_Sgn)
							Call Func_Do_Action(Obj_Wbtn_Sgn,"TRUE","")
							
							Call Func_Object_Exists(Obj_Lnk_Chg_Email)
							Call Func_Do_Action(Obj_Lnk_Chg_Email,"TRUE","")
							
'							obj_pg_Home.Sync
							
							If Obj_Wbele_chg_mail.exist Then								
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_chg_mail.getroproperty("innertext") &" displayed sucessfully ","Pass",now	
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName," Failed to display the chnage E-mail page. Please check again. " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail: Login failed "
								
							End If
							
							Call Func_Object_Exists(Obj_Wbedt_Usr_ID)
							Call Func_Do_Action(Obj_Wbedt_Usr_ID,Environment("User"),WebEdit)
															
							Call Func_Object_Exists(Obj_Wbedt_Usr_pwd)
							Call Func_Do_Action(Obj_Wbedt_Usr_pwd,Environment("Pwd"),WebEdit)

							Call Func_Object_Exists(Obj_Wbedt_email)
							Call Func_Do_Action(Obj_Wbedt_email,Environment("New_User_ID"),WebEdit)
							
							Call Func_Object_Exists(Obj_Lnk_Sbmt)
							Call Func_Do_Action(Obj_Lnk_Sbmt,"TRUE","")
							
							If Obj_Lnk_chg_err.exist(3) Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_Lnk_chg_err.getroproperty("innertext")&" Change the testdata and try again. " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail: Login failed "
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,Obj_Wbele_chg_Suc.getroproperty("innertext") &" displayed from webpage ","Pass",now							
							End If
							
							ExecuteFile(strControlResubalePath)
							
							Call Func_Object_Exists(Obj_Lnk_Chg_Email)
							Call Func_Do_Action(Obj_Lnk_Chg_Email,"TRUE","")
							
							Call Func_Object_Exists(Obj_Wbedt_Usr_ID)
							Call Func_Do_Action(Obj_Wbedt_Usr_ID,Environment("New_User_ID"),WebEdit)
															
							Call Func_Object_Exists(Obj_Wbedt_Usr_pwd)
							Call Func_Do_Action(Obj_Wbedt_Usr_pwd,Environment("Pwd"),WebEdit)

							Call Func_Object_Exists(Obj_Wbedt_email)
							Call Func_Do_Action(Obj_Wbedt_email,Environment("User"),WebEdit)
							
							Call Func_Object_Exists(Obj_Lnk_Sbmt)
							Call Func_Do_Action(Obj_Lnk_Sbmt,"TRUE","")
	
						next

				end if

'#################################### End of Test Case Validation #########################################################################################

'''***********************Writting final Summary Result for the test case in the summary report file & Uploading the Detailed test execution report to Quality Center Detailed result folder*********
Call Func_WriteSummaryReportAndUploadToQC()

'************************End of Summary Report and Uploading detailed report into Quality Center******************************************************************************************************

''****************************************************Script End***************************************************************************************************************************************

'Set obj_pg_Home=Browser("CreationTime:=0").Page("Micclass:=Page")
'
'Set Wbele_chg_mail=Description.Create
'Wbele_chg_mail("innertext").Value="Software Passport Change E-mail Address"
'Wbele_chg_mail("html tag").Value="H1"
'Set Obj_Wbele_chg_mail=obj_pg_Home.WebElement(Wbele_chg_mail)
'Obj_Wbele_chg_mail.Highlight
'Obj_Wbele_chg_mail.
'
'Set Lnk_Chg_Email=Description.Create
'Lnk_Chg_Email("name").Value="Change Email ID"
'Lnk_Chg_Email("html tag").Value="A"
'Set Obj_Lnk_Chg_Email=obj_pg_Home.Link(Lnk_Chg_Email)
'Obj_Lnk_Chg_Email.Highlight
'
'Set Wbedt_Usr_ID=Description.Create
'Wbedt_Usr_ID("name").Value="userId"
'Wbedt_Usr_ID("html tag").Value="INPUT"
'Set Obj_Wbedt_Usr_ID=obj_pg_Home.WebEdit(Wbedt_Usr_ID)
'Obj_Wbedt_Usr_ID.Highlight
'
'Set Wbedt_Usr_pwd=Description.Create
'Wbedt_Usr_pwd("name").Value="password"
'Wbedt_Usr_pwd("html tag").Value="INPUT"
'Set Obj_Wbedt_Usr_pwd=obj_pg_Home.WebEdit(Wbedt_Usr_pwd)
'
'Set Lnk_Sbmt=Description.Create
'Lnk_Sbmt("name").Value="Submit "
'Lnk_Sbmt("html tag").Value="A"
'Set Obj_Lnk_Sbmt=obj_pg_Home.Link(Lnk_Sbmt)
'
'Set Lnk_chg_err=Description.Create
'Lnk_chg_err("name").Value="This e-mail address is already associated.*"
'Lnk_chg_err("html tag").Value="A"
'Set Obj_Lnk_chg_err=obj_pg_Home.Link(Lnk_chg_err)
'
'Set Wbele_chg_Suc=Description.Create
'Wbele_chg_Suc("innertext").Value="Your HP Passport account is modified successfully."
'Wbele_chg_Suc("html tag").Value="DIV"
'Set Obj_Wbele_chg_Suc=obj_pg_Home.WebElement(Wbele_chg_Suc)
'
