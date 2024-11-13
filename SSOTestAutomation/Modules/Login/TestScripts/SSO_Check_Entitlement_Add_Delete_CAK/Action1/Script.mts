'###############################################################################################################
'Test Script Name                          :  SSO_Check_Entitlement_Add_Delete_CAK
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Check_Entitlement_Add_Delete_CAK
'Test Case Name                            :  SSO_Check_Entitlement_Add_Delete_CAK
'Author                                    :  Anil Kumar V
'Designed Date                             :  10 Feb 2015
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="SSO_Check_Entitlement_Add_Delete_CAK"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="SSO_Check_Entitlement_Add_Delete_CAK"
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
                
                DataTable.ImportSheet environment("strDataFile"), "OSP_Check_EntitlementCAK", "Global"
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
							Environment("SAID")=DataTable.GetSheet("Global").GetParameter("SAID").ValueByRow(i)
							
							
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
							
							Call Func_Object_Exists(Obj_lnk_chk_entl)			
							Call Func_Do_Action(Obj_lnk_chk_entl,"TRUE","")
'							Obj_Lnk_Edit_Dshbrd.Click
							Browser("CreationTime:=0").Sync			
							If Obj_wbele_chk_entl.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_wbele_chk_entl.GetRoproperty("innertext")&" page displayed sucessfully" ,"Pass",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Check Entitlement Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If
													
							SAID_val=Split(Environment("SAID"),",")
														
							If ubound(SAID_val)<1 Then
								Call Func_Object_Exists(Obj_Wbet_SAID)
								Call Func_Do_Action(Obj_Wbet_SAID,Environment("SAID"),WebEdit)
								
								Call Func_Object_Exists(obj_Wbtn_add)			
								Call Func_Do_Action(obj_Wbtn_add,"TRUE","")
								Browser("CreationTime:=0").Sync
								If obj_Wbtbl_cnt.Exist Then
									Call Func_Object_Exists(obj_Wbtbl_cnt)
									Row_No=obj_Wbtbl_cnt.GetRowWithCellText(Environment("SAID"))
									If Row_No>0 Then
										Func_CreateRow strDtlRptFileName,strTestCaseName, "CAK sucessfully added into the list " & obj_Wbtbl_cnt.Getcelldata(Row_No,1),"Pass",now	
										else
										Func_CreateRow strDtlRptFileName,strTestCaseName, "CAK not added into the list","Fail",now	
									End If
								End If
								else
								
							For k = 0 To ubound(SAID_val) Step 1
									Obj_Wbet_SAID.Refreshobject
									Call Func_Object_Exists(Obj_Wbet_SAID)
									Call Func_Do_Action(Obj_Wbet_SAID,SAID_val(k),WebEdit)
									
									obj_Wbtn_add.Refreshobject
									Call Func_Object_Exists(obj_Wbtn_add)			
									Call Func_Do_Action(obj_Wbtn_add,"TRUE","")
									Browser("CreationTime:=0").Sync
									obj_Wbtbl_cnt.Refreshobject
									If obj_Wbtbl_cnt.Exist Then
										Call Func_Object_Exists(obj_Wbtbl_cnt)
										Row_No=obj_Wbtbl_cnt.GetRowWithCellText(SAID_val(k))
										If Row_No>0 Then
											Func_CreateRow strDtlRptFileName,strTestCaseName, "CAK sucessfully added into the list " & obj_Wbtbl_cnt.Getcelldata(Row_No,1),"Pass",now	
											else
											Func_CreateRow strDtlRptFileName,strTestCaseName, "CAK is Invalid" &SAID_val(k),"Pass",now
										End If
									End If
								Next
							
							End If
							
								
'							If obj_Wbtbl_cnt.exist Then
'								For del = 1 To 0 Step 1 
'									row_del=obj_Wbtbl_cnt.GetRowWithCellText(SAID_val(del))
'									obj_Wbtbl_cnt.ChildItem(row_del,4,"WebButton",0).click
'									wait(1)
'								Next
'							End If
							
'							ExecuteFile(strControlResubalePath)	
'							obj_Wbtbl_cnt.Refreshobject
'							If obj_Wbtbl_cnt.exist Then
'								For avail = 2 To obj_Wbtbl_cnt.rowcount Step 1 
'									Func_CreateRow strDtlRptFileName,strTestCaseName, left(obj_Wbtbl_cnt.Getcelldata(avail,1),12)&" CAK/SAID available in list","Pass",now
'								Next
'							End If
																																			
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

