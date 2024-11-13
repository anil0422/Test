'###############################################################################################################
'Test Script Name                          :  OSP_Manage_Entitlements
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  OSP_Manage_Entitlements
'Test Case Name                            :  OSP_Manage_Entitlements
'Author                                    :  Anil Kumar V
'Designed Date                             :  20 May 2015
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="OSP_Manage_Entitlements"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="OSP_Manage_Entitlements"
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

				'Invoke Browser and Launch application
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" Then
                    Call Func_InvokeBrowser(Environment("Browser_IE"),Environment("URL"))
                Else
                    Call Func_InvokeBrowser(Environment("Browser"),Environment("URL"))
                End If

                If environment("flag")=1 Then
					strDataFile = strDataFile_QC
					environment("strDataFile")=strDataFile
				End If

'#################################### TEST  CASE VALIDATION STARTS HERE ##################################################################################
                
                DataTable.ImportSheet environment("strDataFile"), "OSP_Manage_Entitlements", "Global"
                environment("row_count") = DataTable.GetSheet("Global").GetRowCount
				
				If environment("row_count") = 0 Then
						Func_CreateRow strDtlRptFileName,environment("strTestCaseName"),"Row count - Data table row count is zero or not exists, please input valid data in the data table","Fail",now
						strSumResultStatus="FAIL"
						Call Func_WriteSummaryReportAndUploadToQC()
						Reporter.ReportEvent micFail, "Input Data", "Row count - Data table row count is zero or not exists, please input valid data in the data table"
						ExitAction "Fail: Row count - Data table row count is zero or not exists, please input valid data in the data table"
				Else

					For i = 1 to environment("row_count")
                            																	
							Environment("Search_By") = DataTable.GetSheet("Global").GetParameter("Search_By").ValueByRow(i)
							Environment("Id") = DataTable.GetSheet("Global").GetParameter("Id").ValueByRow(i)
							Environment("Emp_ID")=DataTable.GetSheet("Global").GetParameter("Emp_ID").ValueByRow(i)
							Environment("Pwd_1")=DataTable.GetSheet("Global").GetParameter("Pwd_1").ValueByRow(i)
							
													
'							obj_pg_Home.Sync
							Do Until Obj_Wbtn_Sgn.exist
							Loop
							
							Call Func_Object_Exists(Obj_Wbtn_Sgn)
							Call Func_Do_Action(Obj_Wbtn_Sgn,"TRUE","")
							
'							obj_pg_Home.Sync
							
							Call Func_Object_Exists(Obj_Wbedt_Un)
							Call Func_Do_Action(Obj_Wbedt_Un,Environment("Emp_ID"),WebEdit)
														
							Call Func_Object_Exists(Obj_Wbedt_Pwd)			
							Call Func_Do_Action(Obj_Wbedt_Pwd,Environment("Pwd_1"),WebEdit)
							
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")
							
							If obj_Wbele_err.exist(3) Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Login failed due to: " & obj_Wbele_err.getroproperty("innertext"),"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Fail: Login failed "
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","Pass",now							
							End If
							
'							Setting.WebPackage("ReplayType")=2
'							Obj_Lnk_Usr.FireEvent "onmouseover"
'							Setting.WebPackage("ReplayType")=1

							Call Func_Object_Exists(Obj_Lnk_mng_entltmnt)			
							Call Func_Do_Action(Obj_Lnk_mng_entltmnt,"TRUE","")
							
							If obj_Wbele_ent_pg.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,obj_Wbele_ent_pg.GetRoproperty("innertext")&" page displayed sucessfully ","Pass",now
							End If
							
							Call Func_Object_Exists(Obj_Wblst_ent_srch_by)			
							Call Func_Do_Action(Obj_Wblst_ent_srch_by,Environment("Search_By"),WebList)
							
							Call Func_Object_Exists(Obj_WbEdt_Ent_id)			
							Call Func_Do_Action(Obj_WbEdt_Ent_id,Environment("Id"),WebEdit)
							
							Call Func_Object_Exists(obj_Wbtn_Ent_srch)			
							Call Func_Do_Action(obj_Wbtn_Ent_srch,"TRUE","")
							
							If obj_Wbele_err_msg.exist(5) then
								Func_CreateRow strDtlRptFileName,strTestCaseName, obj_Wbele_err_msg.GetRoproperty("innertext")& " Please change the data in testdata file " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the data in the file"
							End If  
							
							Call Func_Object_Exists(Obj_Wblst_ent_said)	
							said_cnt=cint(Obj_Wblst_ent_said.GetRoProperty("items count"))						
							If said_cnt<=0 Then
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Entitlement count is less than or equal to zero. Please check test data.","Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the data in the file"
							End If
																			
'							ExecuteFile(strControlResubalePath)							
'							Setting.WebPackage("ReplayType")=2
'							Obj_Lnk_Usr.FireEvent "onmouseover"
'							Setting.WebPackage("ReplayType")=1
'							Obj_Lnk_Emp_Acc.Refreshobject
							
'							If not Obj_Lnk_Emp_Acc.exist(5) then
'								Func_CreateRow strDtlRptFileName,strTestCaseName," Employee Access Link was removed as its validated  " ,"Pass",now
'								else
'								Func_CreateRow strDtlRptFileName,strTestCaseName," Employee Access Link was not removed as its validated  " ,"Fail",now
'							End IF
																												
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
'Set obj_pg_Home=Browser("CreationTime:=0").Page("Micclass:=Page")
'
'Set Wblst_ent_said=Description.Create
'Wblst_ent_said("name").Value="_entitlementmanagement_WAR_hpospportlet_saidList"
'Set Obj_Wblst_ent_said=obj_pg_Home.WebList(Wblst_ent_said)
'Obj_Wblst_ent_said.highlight

'
