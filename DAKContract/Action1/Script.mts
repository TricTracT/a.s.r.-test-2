Option Explicit

Dim vboolDone
Dim vstrEnvironment
Dim vstrTempDir

Print "Begin DAKContract " & Now
''''''''''''''''''''''''''''''
' Start DAK based upon DataTable. For each row, proces it through DAK and compare expected result with actual result
vstrEnvironment = "O"
vstrTempDir = "c:\data\DAK\"

'Call subCleanDAKBeforeTest(vstrEnvironment)
Call subDeleteDirectoryContents(vstrTempDir)

'Clean database
Call subCleanDAKBeforeTest(vstrEnvironment, vstrTempDir)

' Inject SQL for testprerequisites
Call subInjectPreRequisites(vstrEnvironment, vstrTempDir)

Call subPrepareDAKInputFromDSV2(vstrEnvironment, "Contracten")

' Start DAK batch which wil get all the copied files to be processed
Call subStartDAKBatchV2(vstrEnvironment, vstrTempDir, "start")

' Wait until batchjob is done. Max 5 seconds with checks with an interval of 1 second
vboolDone = fncWaitUntilJobDoneV2(vstrEnvironment, vstrTempDir, 5, 1)

If vboolDone Then
	Call subVerifyDAKOutputV2(vstrEnvironment, "Contract")	
End If

Reporter.ReportEvent micPass,"DAKContract", "DAKContract uitgevoerd"
Print "Einde DAKContract " & Now
'''''''''''''''''''''''''''''' 

Sub subInjectPreRequisites(vstrEnvironment, vstrTempDir)

	Dim vstrSQL
	Dim vstrFileName
	Dim vstrResult
	 
	vstrFileName = "ContractTestPrerequisites.sql"
	
	'Get request contracts from ALM
	call subGetResourceUFT(vstrFileName, vstrTempDir)
	
	If vstrFileName = "" Then
		Reporter.ReportEvent micFail, "subInjectPreRequisites", "SQL file not loaded from QC with filename: " & vstrFileName
	Else
		vstrSQL = fncReadEntireFile(vstrTempDir & vstrFileName)
	End If
	vstrSQL = fncReadEntireFile(vstrTempDir & vstrFileName)
	
	vstrResult = fncRunSQLStatementOnDBV2(vstrEnvironment, vstrTempDir, vstrSQL)	
	
End Sub


