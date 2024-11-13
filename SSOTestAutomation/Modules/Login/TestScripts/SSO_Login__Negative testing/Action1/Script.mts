'###############################################################################################################
'Test Script Name                          :  OSP_Login__Negative testing
'Test Objective/Description         	   :  To Login and Logout into OSP Application
'Test Case ID                              :  OSP_Login__Negative testing
'Test Case Name                            :  OSP_Login__Negative testing
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
                environment("StrTestCaseID")="OSP_Login__Negative testing"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="OSP_Login__Negative testing"
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
               Environment("Browser_chrome") = "C:\Users\"&User&"\AppData\Local\Google\Chrome\Application\chrome.exe"
				'Invoke Browser and Launch application
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" Then
                    Call Func_InvokeBrowser(Environment("Browser_IE"),Environment("URL"))
                    'Call Func_InvokeBrowser(Environment("Browser_FF"),Environment("URL"))
                    'Call Func_InvokeBrowser(Environment("Browser_chrome"),Environment("URL"))
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
                            																	
'							environment("User") = DataTable.GetSheet("Global").GetParameter("User").ValueByRow(i)
'							environment("Pwd") = DataTable.GetSheet("Global").GetParameter("Pwd").ValueByRow(i)
							
							obj_pg_Home.Sync
							
							Call Func_Object_Exists(Obj_Lnk_Sgn)
							Call Func_Do_Action(Obj_Lnk_Sgn,"TRUE","")
							
							obj_pg_Home.Sync
							wait 2
							Call Func_Object_Exists(Obj_Wbedt_Un)
							Call Func_Do_Action(Obj_Wbedt_Un,Environment("User"),WebEdit)
							
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")
							
							If obj_Wbele_err.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Error Message displayed as " & obj_Wbele_err.getroproperty("innertext"),"Pass",now	
							End IF
							Browser("CreationTime:=0").Refresh
							
							Call Func_Object_Exists(Obj_Wbedt_Pwd)			
							Call Func_Do_Action(Obj_Wbedt_Pwd,Environment("Pwd"),WebEdit)
							
							Obj_Lnk_USgn.Refreshobject
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")
							
							obj_Wbele_err.Refreshobject
							If obj_Wbele_err.exist Then
								Func_CreateRow strDtlRptFileName,strTestCaseName, "Error Message displayed as " & obj_Wbele_err.getroproperty("innertext"),"Pass",now	
							End IF
							
							Browser("CreationTime:=0").Refresh
							
							Obj_Wbedt_Un.Refreshobject
							Call Func_Object_Exists(Obj_Wbedt_Un)
							Call Func_Do_Action(Obj_Wbedt_Un,Environment("User"),WebEdit)
							
							Obj_Wbedt_Pwd.Refreshobject
							Call Func_Object_Exists(Obj_Wbedt_Pwd)			
							Call Func_Do_Action(Obj_Wbedt_Pwd,Environment("Pwd"),WebEdit)
							
							Obj_Lnk_USgn.Refreshobject
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")
'							
'							obj_Wbele_err.Refreshobject
'							If obj_Wbele_err.exist Then
'								Func_CreateRow strDtlRptFileName,strTestCaseName, "Login failed due to: " & obj_Wbele_err.getroproperty("innertext"),"Fail",now
'								Call Func_WriteSummaryReportAndUploadToQC()
'								ExitAction "Fail: Login failed "
'								else
								Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","Pass",now							
'							End If
							

'							obj_pg_Home.RefreshObject
'							obj_pg_Home.Sync
							
							Do Until Obj_Lnk_Usr.exist
							Loop
							
'							Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","Pass",now
							
							Setting.WebPackage("ReplayType")=2
							Obj_Lnk_Usr.FireEvent "onmouseover"
							Setting.WebPackage("ReplayType")=1			
							

							
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
'Set obj_pg=Browser("CreationTime:=0").Page("index:=0")
''obj_pg.highlight
'
'Set cart=Description.Create
'cart("html tag").Value="DIV"
'cart("index").Value=48
'Set obj_cart=obj_pg.WebElement(cart)
'obj_cart.Highlight
'Setting.webpackage("ReplayType")=2
'obj_cart.FireEvent "onmouseover"
'Setting.webpackage("ReplayType")=1
'
'wait(2)
'
'Set img_sel=Description.Create
''img_sel("file name").Value="logo_9919466_80_40.gif"
'img_sel("alt").Value="Best Buy" 'hpshopping.com
'img_sel("html tag").Value="IMG"
'img_sel("image type").Value="Plain Image"
'Set obj_img_sel=obj_pg.Image(img_sel)
'obj_img_sel.highlight
'obj_img_sel.Click
