'###############################################################################################################
'Test Script Name                          :  SSO_Configuration_Items
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  SSO_Configuration_Items
'Test Case Name                            :  SSO_Configuration_Items
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
                environment("StrTestCaseID")="OSP_Configuration_Items"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="OSP_Configuration_Items"
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
                
                DataTable.ImportSheet environment("strDataFile"), "OSP_Configuration_Items", "Global"
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
							Environment("Configuration_Name")=DataTable.GetSheet("Global").GetParameter("Configuration_Name").ValueByRow(i)
							Environment("Product_Type")=DataTable.GetSheet("Global").GetParameter("Product_Type").ValueByRow(i)
							Environment("Status")=DataTable.GetSheet("Global").GetParameter("Status").ValueByRow(i)
							Environment("Create_New_Environment")=DataTable.GetSheet("Global").GetParameter("Create_New_Environment").ValueByRow(i)
							Environment("Environment_Name")=DataTable.GetSheet("Global").GetParameter("Environment_Name").ValueByRow(i)
							Environment("SAID")=DataTable.GetSheet("Global").GetParameter("SAID").ValueByRow(i)
							Environment("Instance")=DataTable.GetSheet("Global").GetParameter("Instance").ValueByRow(i)
							Environment("Product_list")=DataTable.GetSheet("Global").GetParameter("Product_list").ValueByRow(i)
							Environment("Description")=DataTable.GetSheet("Global").GetParameter("Description").ValueByRow(i)
							Environment("Product_Version")=DataTable.GetSheet("Global").GetParameter("Product_Version").ValueByRow(i) 
							Environment("Operating_System")=DataTable.GetSheet("Global").GetParameter("Operating_System").ValueByRow(i)
							
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
																
							Call Func_Object_Exists(Obj_Lnk_Cnfg_itms)			
							Call Func_Do_Action(Obj_Lnk_Cnfg_itms,"TRUE","")
							
							wait(20)
							If Obj_Wbele_Cnfg_Itms.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, Obj_Wbele_Cnfg_Itms.GetRoproperty("innertext")&" page displayed sucessfully" ,"Pass",now
								else
								Func_CreateRow strDtlRptFileName,strTestCaseName,"Failed to display Configuration Items Page " ,"Fail",now
								Call Func_WriteSummaryReportAndUploadToQC()
								ExitAction "Failed to display the Page"
							End If

							Call Func_Object_Exists(Obj_Lnk_nCnfg_itms)			
							Call Func_Do_Action(Obj_Lnk_nCnfg_itms,"TRUE","")
																		
							Call Func_Object_Exists(Obj_Wbedt_cnfg_name)			
							Call Func_Do_Action(Obj_Wbedt_cnfg_name,Environment("Configuration_Name"),WebEdit)
							
							Call Func_Object_Exists(Obj_Wbele_Act_Ys)			
							Call Func_Do_Action(Obj_Wbele_Act_Ys,"TRUE","")
							
							If Environment("Product_Type")<>"" Then
								Call Func_Object_Exists(obj_Wblst_Prd_nm)			
								Call Func_Do_Action(obj_Wblst_Prd_nm,Environment("Product_Type"),WebList)
							End If
							
							If Environment("Status")<>"" Then
								Call Func_Object_Exists(obj_Wblst_Prd_st)			
								Call Func_Do_Action(obj_Wblst_Prd_st,Environment("Status"),WebList)
							End If							
							
							If Environment("Create_New_Environment")="Yes" Then
								Call Func_Object_Exists(Obj_Lnk_new_Env)			
								Call Func_Do_Action(Obj_Lnk_new_Env,"TRUE","")
								
								Call Func_Object_Exists(Obj_Wbedt_new_env_name)
								Call Func_Do_Action(Obj_Wbedt_new_env_name,Environment("Environment_Name"),WebEdit)
								
								Call Func_Object_Exists(obj_Wblst_env_SAID)			
								Call Func_Do_Action(obj_Wblst_env_SAID,Environment("SAID"),WebList)

								Call Func_Object_Exists(obj_Wblst_env_typ)			
								Call Func_Do_Action(obj_Wblst_env_typ,Environment("Instance"),WebList)
								
								Call Func_Object_Exists(Obj_Wbtn_Sbmt)			
								Call Func_Do_Action(Obj_Wbtn_Sbmt,"TRUE","")
								
								
								Call Func_Object_Exists(Obj_Wbtn_Cancel)			
								Call Func_Do_Action(Obj_Wbtn_Cancel,"TRUE","")
							End If
							
							If Environment("Environment_Name")<>"" Then								
								Call Func_Object_Exists(obj_Wblst_env_nam)			
								Call Func_Do_Action(obj_Wblst_env_nam,Environment("Environment_Name"),WebList)
							End If
							
							wait(2)
							If Environment("Product_list")<>"" Then								
								Call Func_Object_Exists(obj_Wblst_Prd_lst)	
								obj_Wblst_Prd_lst.WaitProperty "disabled","0"								
								Call Func_Do_Action(obj_Wblst_Prd_lst,Environment("Product_list"),WebList)
							End If
							
							If Environment("Description")<>"" Then
								Call Func_Object_Exists(Obj_Wbedt_Cnfg_Desc)			
								Call Func_Do_Action(Obj_Wbedt_Cnfg_Desc,Environment("Description"),WebList)
							End If
							
							Call Func_Object_Exists(obj_Wblst_Prd_ver)	
							obj_Wblst_Prd_ver.WaitProperty "default value","Select"								
							Call Func_Do_Action(obj_Wblst_Prd_ver,Environment("Product_Version"),WebList)
							
							Call Func_Object_Exists(obj_Wblst_OS)
							obj_Wblst_OS.WaitProperty "default value","Select"								
							Call Func_Do_Action(obj_Wblst_OS,Environment("Operating_System"),WebList)
							
							Call Func_Object_Exists(Obj_Wbtn_Cnf_Sbmt)			
							Call Func_Do_Action(Obj_Wbtn_Cnf_Sbmt,"TRUE","")
							
							wait(10)
							If Obj_Cnfig_suc.exist then
								Data_Val=wbtbl_Cng_Itm.GetRowWithCellText(Environment("Configuration_Name"))
									Func_CreateRow strDtlRptFileName,strTestCaseName," New Configuration item created Sucessfully ","Done",now								
							End IF							
																		
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
'Set Wbtn_Cnf_Sbmt=Description.Create
'Wbtn_Cnf_Sbmt("html id").Value="createConfigItemBtn"
'Wbtn_Cnf_Sbmt("name").Value="Submit"
'Wbtn_Cnf_Sbmt("html tag").Value="BUTTON"
'Set Obj_Wbtn_Cnf_Sbmt=obj_pg_Home.Webbutton(Wbtn_Cnf_Sbmt)
'Obj_Wbtn_Cnf_Sbmt.Highlight
'Obj_Wbtn_Cnf_Sbmt.Click
