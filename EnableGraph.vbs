
Path = "E:\APIAutomation\Project\Results.xlsx"
    ' Create a new instance of Excel and make it visible.
    Set oXL = CreateObject("Excel.Application")
    oXL.Visible = True

    ' Add a new workbook and set a reference to Sheet1.
    Set oBook = oXL.Workbooks.Open(Path)
    'Set oSheet = oBook.Sheets(1)

'Get the System Name through Windows shell script

Set wshShell = CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
'msgbox"User Name: " & strUserName

    'Import previously created BAS module file
    oXL.VBE.ActiveVBProject.VBComponents.Import "E:\APIAutomation\Project\Macro.bas"



    ' Now run the macro, passing oSheet as the first parameter
    oXL.Run "Fact"
	oBook.save
	oXL.Quit
	

    'oXL.UserControl = True
    Set oSheet = Nothing
    Set oBook = Nothing
    Set oXL = Nothing
