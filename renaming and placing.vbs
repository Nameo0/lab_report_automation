'Created on Jan. 04, 2016
'By Leon Stanislaw Kozlowski
'
'Will create a lab report with th curren date
'in the proper folde once the script is run.
'It will not over-write other lab reports due
'to checking if a lab report for today exists.


'BEGIN INITILIZATION AND SETTING THINGS UP

dim objFSO	'As FileSystemObject
'dim objShell	'As ObjectShell
dim scrShell	'As ScriptShell
dim objWord	'As Application
dim curDirectory
dim date_year
dim date_month
dim date_day
dim templateToUse
dim theFileName
dim OFFICE_PATH

set objFSO	= CreateObject("Scripting.FileSystemObject")
set objWord	= CreateObject("Word.Application")
'set objShell	= wscript.CreateObject("Shell.Application")

'Used to change working directory
curDirectory = "L:\Lab Reports"
set scrShell = CreateObject("Wscript.Shell")
scrShell.CurrentDirectory = curDirectory

date_year	= Year(Now)
'Converts YYYY to YY
date_year_short = Right(date_year,2)
'Makes sure all month values are 2 digit
date_month	        = Right("0" & Month(Now), 2)
date_month_name	= MonthName(date_month)
'Makes sure all day values are 2 digit
date_day 	= Right("0" & Day(Now), 2)

'END OF INITIALIZATION AND SETTING THINGS UP

'Will return TRUE if the year does not exist and will create a folder with the year's name
If Not(objFSO.FolderExists(date_year)) Then
	'Creates folder with current year
	objFSO.CreateFolder(date_year)

	scrShell.CurrentDirectory = curDirectory & "\" & date_year
	objFSO.CreateFolder("USU")
End If

'Changes working directory
scrShell.CurrentDirectory = curDirectory & "\" & date_year & "\USU"
curDirectory = scrShell.CurrentDirectory

'Will return TRUE if the month does not exist and will create a folder with the month's name
If Not(objFSO.FolderExists(date_month_name)) Then
	'Creates folder with current month
	objFSO.CreateFolder(date_month_name)
End If

'Changes working directory
scrShell.CurrentDirectory = curDirectory & "\" & date_month_name
curDirectory = scrShell.CurrentDirectory

'Checks if a lab report exists for today, if not one is copied and made
templatetoUse = "CLT - Report Form - USU (Break Hours).docx" 'Make sure that the '.docs' part exists when you input the new file template
theFileName = "USU Lab Report " & date_month & "-" & date_day & "-" & date_year_short & ".docx"
If Not(objFSO.FileExists(curDirectory & "\" & theFileName)) Then
	objFSO.CopyFile("L:\Lab Reports\" & templateToUse, curDirectory & "\")
	objFSO.MoveFile("CLT - Report Form - USU (Break Hours).docx", theFileName)
Else
	MsgBox("The file alread exists")
End If

objWord.Visible = TRUE
objWord.Documents.Open("L:\Lab Reports\2016\USU\January\USU Lab Report " & date_month & "-" & date_day & "-" & date_year_short & ".docx")

'objShell.Open(theFileName)

'MsgBox("tacocat")