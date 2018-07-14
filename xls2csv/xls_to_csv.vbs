csv_format = 6

Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim oExcel
Set oExcel = CreateObject("Excel.Application")

Dim oBook
Dim oColumnColor
If WScript.Arguments.Count < 1 Then
	Set oFolder = objFSO.GetFolder(".")
	Set oFiles = oFolder.Files
	For Each oFile In oFiles
		If Right(oFile.Path,4)=".xls" Then
			Set oBook = oExcel.Workbooks.Open(oFile.Path)
			For i=1 to oBook.ActiveSheet.Columns.Count
				If oBook.ActiveSheet.Cells(1,i).Value = "" Then
					Exit For
				End If
				oColumnColor = oBook.ActiveSheet.Cells(1,i).Interior.Color
				If oBook.ActiveSheet.Columns(i).Interior.ColorIndex Then
				Else
					oBook.ActiveSheet.Columns(i).Delete
					i = i-1
				End If
			Next
			dest_file = objFSO.GetAbsolutePathName(".\csv\" & Replace(objFSO.GetFileName(oFile.Path), "xls", "csv"))
			'Msgbox dest_file
			oBook.SaveAs dest_file, csv_format
			oBook.Close False
		End If
	Next
	oExcel.Quit
Else
	Set oFile = objFSO.GetFile(Wscript.Arguments.Item(0))
	If Right(oFile.Path,4)=".xls" Then
		Set oBook = oExcel.Workbooks.Open(oFile.Path)
		For i=1 to oBook.ActiveSheet.Columns.Count
			If oBook.ActiveSheet.Cells(1,i).Value = "" Then
				Exit For
			End If
			oColumnColor = oBook.ActiveSheet.Cells(1,i).Interior.Color
			If oBook.ActiveSheet.Columns(i).Interior.ColorIndex Then
			Else
				oBook.ActiveSheet.Columns(i).Delete
				i = i-1
			End If
		Next
        dest_file = objFSO.GetAbsolutePathName(".\csv\" & Replace(objFSO.GetFileName(oFile.Path), "xls", "csv"))
        'Msgbox dest_file
        oBook.SaveAs dest_file, csv_format
		oBook.Close False
	End If
	oExcel.Quit
End If

