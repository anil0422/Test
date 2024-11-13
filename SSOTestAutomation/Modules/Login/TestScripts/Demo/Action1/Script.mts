'###############################################################################################################
'Test Script Name                          :  Login
'Test Objective/Description         	   :  Launch Application
'Test Case ID                              :  Demo
'Test Case Name                            :  DEMO
'Author                                    : 
'Designed Date                             :  1 July.2013
'Last Modified By                          :  
'Last Modified Date                        :  
'###############################################################################################################

'********************************************************Variable Declaration******************************************************************************

				Dim strTestcaseID,strModuleName,strTestCaseName,strSumResultStatus,strFileName,strEndTime,strDuration
				Dim oQTP, strInputPath, strProjectName

'************************************************************************End of Variable declaration *****************************************************

'****************************************Setting the Project Name in the environment variable*******************************************************
                StrTestCaseID= environment("TestName")
                environment("StrTestCaseID")="Demo"
                strModuleName= environment("strModuleName")				

                strTestCaseName= environment("TestName")
                environment("strTestCaseName")="Demo"
                StrStartTime=now
                StrBeginTimer=Timer

                strProjectName = environment("ProjectName")
                environment("strOSName") = Environment.Value("OS")
                
                If environment("strOSName") = "Microsoft Windows 7 Workstation" Or environment("strOSName") = "Microsoft Windows Vista Workstation" Or environment("strOSName") = "Microsoft Windows Vista Server" Then
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
					qtpApp.Folders.Add ("[QualityCenter] Subject\C4\Automation Scripts\" & environment("ProjectName") & "\DriverData"),1
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
                    Call Func_InvokeBrowser(Environment("Browser_x86"),Environment("URL"))
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
                            																	
							environment("User") = DataTable.GetSheet("Global").GetParameter("User").ValueByRow(i)
							environment("Pwd") = DataTable.GetSheet("Global").GetParameter("Pwd").ValueByRow(i)
							
							obj_pg_Home.Sync
							
							Call Func_Object_Exists(Obj_Lnk_Sgn)
							Call Func_Do_Action(Obj_Lnk_Sgn,"TRUE","")
							
							obj_pg_Home.Sync
							wait 2
							Call Func_Object_Exists(Obj_Wbedt_Un)
							Call Func_Do_Action(Obj_Wbedt_Un,Environment("User"),WebEdit)
														
							Call Func_Object_Exists(Obj_Wbedt_Pwd)			
							Call Func_Do_Action(Obj_Wbedt_Pwd,Environment("Pwd"),WebEdit)
							
							Call Func_Object_Exists(Obj_Lnk_USgn)
							Call Func_Do_Action(Obj_Lnk_USgn,"TRUE","")
							
'							obj_pg_Home.RefreshObject
'							obj_pg_Home.Sync
							
							Do Until Obj_Lnk_Usr.exist
							Loop
							
							Func_CreateRow strDtlRptFileName,strTestCaseName," User Sucessfully Logged into OSP Application","Pass",now
							
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



str="malayalam"
'strrev=strreverse(str)
a=""
strlen=Len(str)
For i = strlen To 1 Step -1
	'a=mid(str,i,1)&a
	a=a&right(left(str,i),1)
Next
	MsgBox a

If str=a Then
	MsgBox "Its a Palindrome"
	else
	MsgBox "Not a Palindrome"
End If

a=1
b=2
c=a+b
MsgBox c


val=48
temp=0
For i = 2 To val/2 Step 1
	If val mod i=0 Then
		msgbox "Not a Prime"
		temp=1
		Exit for
	End If
Next
If temp=0 Then
	MsgBox "Prime"
End If



StrBody = "Hello"&vbcrlf&vbcrlf&"Please find attached your Domain Account details for Micro Focus."&vbcrlf&vbcrlf&"Additional Documents -"&vbcrlf&vbcrlf&"1) The Cornwall Program document outlines next steps to log into MF and change your password."&vbcrlf&"2) New Contractor CEID request outlines the UAM request process."&vbcrlf&vbcrlf&"Please follow the documents to setup your MF account access."&vbcrlf&vbcrlf&"Please Note: This is auto generated email, Do-NOT Reply.  Please contact Helpdesk @ (https://csc.service-now.com/selfservice/) for any Support."&vbcrlf&vbcrlf&"Regards"

Msgbox "<b>"&StrBody&"</b>
